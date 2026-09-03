import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st

from shared import (
    APP_VERSION,
    REQUIRED_FILES,
    OPTIONAL_FILES,
    OPTIONAL_FILE_LABELS,
    DEBTOR_EXPORT_FILE,
    load_metadata,
    publish_files,
    save_metadata,
    upload_file,
    validate_feestdagen_xlsx,
    validate_required_files_in_folder,
    validate_debtor_export_xlsx,
    render_environment_banner,
    get_page_title,
    is_test_environment,
    get_locatie_naam,
    get_locatie_logo,
    get_beheer_wachtwoord,
    compute_otif_from_folder,
    save_daily_snapshot,
)
from pathlib import Path
import tempfile

st.set_page_config(layout="wide", page_title=get_page_title(f"Capaciteitsplanning {get_locatie_naam()} - Beheer"))

_logo = get_locatie_logo()
if os.path.exists(_logo):
    st.sidebar.image(_logo, use_container_width=True)
st.sidebar.caption(APP_VERSION)
render_environment_banner("Beheer")

password = st.sidebar.text_input("Wachtwoord", type="password")
if password != get_beheer_wachtwoord():
    st.stop()

st.title(f"Capaciteitsplanning {get_locatie_naam()} – Beheer")
st.caption("Upload per bestand en publiceer daarna de volledige dataset voor alle kijkers.")
if is_test_environment():
    st.warning("Let op: dit is de TEST-beheeromgeving. Controleer de bucket/secrets en publiceer alleen bewust bestanden. Als Test dezelfde bucket gebruikt als Main, schrijf je naar dezelfde dataopslag.")

# ── Laatste publicatie ─────────────────────────────────────────────────────────
meta = load_metadata()
top1, top2 = st.columns([1, 2])

with top1:
    st.subheader("Laatste publicatie")
    if meta:
        st.write(f'**Laatste update:** {meta.get("published_at", "-")}')
        st.write(f'**Door:** {meta.get("published_by", "-")}')
        st.write(f'**Toelichting:** {meta.get("notes", "-")}')
    else:
        st.info("Nog geen dataset gepubliceerd.")

with top2:
    st.subheader("Vereiste bestanden")
    st.write(", ".join(REQUIRED_FILES))
    st.subheader("Optionele bestanden")
    st.write(", ".join(OPTIONAL_FILES))

st.markdown("---")

# ── Bestanden uploaden ─────────────────────────────────────────────────────────
st.subheader("1. Upload bestanden")
st.caption(
    "Upload elk bestand in het juiste vak. "
    "Publiceren vervangt alleen de bestanden die je hebt geüpload in deze sessie."
)

uploads: dict[str, st.runtime.uploaded_file_manager.UploadedFile | None] = {}
upload_cols = st.columns(2)
for i, fname in enumerate(REQUIRED_FILES):
    with upload_cols[i % 2]:
        uploads[fname] = st.file_uploader(f"Upload {fname}", type=["xlsx"], key=fname)

st.markdown("**Optionele bestanden**")
opt_cols = st.columns(2)
for i, fname in enumerate(OPTIONAL_FILES):
    with opt_cols[i % 2]:
        uploads[fname] = st.file_uploader(
            OPTIONAL_FILE_LABELS.get(fname, f"Upload {fname} (optioneel)"),
            type=["xlsx"], key=fname
        )

st.markdown("---")

# ── Validatiestatus ────────────────────────────────────────────────────────────
st.subheader("2. Validatie")

# Check what's currently in the cloud
try:
    from shared import list_cloud_files
    cloud_files = set(list_cloud_files())
except Exception:
    cloud_files = set()

validation_rows = []
for fname in REQUIRED_FILES:
    validation_rows.append(
        {
            "Bestand": fname,
            "Vereist": "Ja",
            "Nieuw geüpload": "✅ Ja" if uploads.get(fname) is not None else "–",
            "Staat al in cloud": "✅ Ja" if fname in cloud_files else "❌ Nee",
        }
    )
for fname in OPTIONAL_FILES:
    validation_rows.append(
        {
            "Bestand": fname,
            "Vereist": "Optioneel",
            "Nieuw geüpload": "✅ Ja" if uploads.get(fname) is not None else "–",
            "Staat al in cloud": "✅ Ja" if fname in cloud_files else "–",
        }
    )
st.dataframe(pd.DataFrame(validation_rows), use_container_width=True, hide_index=True)

# ── Publiceren ─────────────────────────────────────────────────────────────────
publisher = st.text_input("Naam beheerder", value="")
notes = st.text_input("Notitie / omschrijving update", value="")
publish = st.button("Publiceer dataset", type="primary", use_container_width=True)

if publish:
    # Collect bytes from newly uploaded files + existing cloud files for validation
    stage: dict[str, bytes] = {}

    # First load existing cloud files for files not newly uploaded
    for fname in REQUIRED_FILES:
        if uploads.get(fname) is not None:
            stage[fname] = uploads[fname].getbuffer().tobytes()
        elif fname in cloud_files:
            try:
                from shared import download_file
                stage[fname] = download_file(fname)
            except Exception as e:
                st.error(f"Kan bestaand bestand '{fname}' niet ophalen: {e}")
                st.stop()

    # Validate in a temp folder
    tmp_dir = Path(tempfile.mkdtemp())
    try:
        for fname, data in stage.items():
            (tmp_dir / fname).write_bytes(data)

        validate_required_files_in_folder(tmp_dir)

        # Upload optional files if provided. Debtor-export wordt gevalideerd omdat
        # deze de segmentkoppeling voor het dashboard voedt.
        for fname in OPTIONAL_FILES:
            if uploads.get(fname) is not None:
                stage[fname] = uploads[fname].getbuffer().tobytes()
                (tmp_dir / fname).write_bytes(stage[fname])
                if fname == DEBTOR_EXPORT_FILE:
                    validate_debtor_export_xlsx(tmp_dir / fname)

        # All valid → upload newly uploaded files to cloud
        newly_uploaded = {
            fname: data
            for fname, data in stage.items()
            if uploads.get(fname) is not None
        }

        if not newly_uploaded:
            st.warning("Je hebt geen nieuwe bestanden geüpload. Er is niets gewijzigd.")
        else:
            with st.spinner("Bestanden uploaden naar cloud…"):
                publish_files(newly_uploaded)

            metadata = {
                "published_at": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
                "published_by": publisher if publisher else "Onbekend",
                "notes": notes if notes else "",
                "files": list(newly_uploaded.keys()),
            }
            save_metadata(metadata)
            st.success(
                f"✅ Dataset succesvol gepubliceerd. "
                f"{len(newly_uploaded)} bestand(en) bijgewerkt: "
                + ", ".join(newly_uploaded.keys())
            )

            # ── OTIF-snapshot op moment van publicatie ─────────────────────
            # De exports bevatten de orderstatus zoals die nu is — dit is het
            # enige betrouwbare moment om de dagelijkse OTIF-waarde vast te
            # leggen. De snapshot wordt slechts één keer per dag opgeslagen
            # (overwrite=False); bij meerdere publicaties op dezelfde dag
            # blijft de eerste (vroegste) snapshot leidend.
            with st.spinner("OTIF-snapshot opslaan…"):
                try:
                    _otif_snap = compute_otif_from_folder(tmp_dir)
                    _snap_saved = save_daily_snapshot(
                        datum=datetime.now().date(),
                        tonnage_plan_kg=None,  # plan-tonnage is niet beschikbaar in manager
                        otif_result=_otif_snap,
                        overwrite=False,
                        backfilled=False,
                    )
                    if _snap_saved and _otif_snap and _otif_snap.get("otif_pct") is not None:
                        st.info(
                            f"📸 OTIF-snapshot opgeslagen: "
                            f"{_otif_snap['otif_pct']:.1f}% "
                            f"({_otif_snap.get('gereed', 0)}/{_otif_snap.get('totaal', 0)} orders op tijd)."
                        )
                    elif not _snap_saved:
                        st.caption("ℹ️ Er bestaat al een OTIF-snapshot voor vandaag — niet overschreven.")
                except Exception as _snap_err:
                    st.warning(
                        f"⚠️ Publicatie geslaagd, maar OTIF-snapshot kon niet worden opgeslagen: "
                        f"{_snap_err}. De snapshot kan handmatig worden aangevuld via de viewer."
                    )
    except Exception as e:
        st.error(f"Publicatie mislukt: {e}")
    finally:
        import shutil
        shutil.rmtree(tmp_dir, ignore_errors=True)

st.markdown("---")

# ── Feestdagen template ────────────────────────────────────────────────────────
st.subheader("3. Template feestdagen.xlsx")
template_df = pd.DataFrame(
    [
        ["2026-01-01", "Nieuwjaarsdag", "Feestdag"],
        ["2026-04-03", "Goede vrijdag", "Feestdag"],
        ["2026-04-06", "2e Paasdag", "Feestdag"],
        ["2026-04-27", "Koningsdag", "Feestdag"],
        ["2026-05-05", "Bevrijdingsdag", "Feestdag"],
        ["2026-05-14", "Hemelvaartsdag", "Feestdag"],
        ["2026-05-25", "2e Pinksterdag", "Feestdag"],
        ["2026-12-25", "1e Kerstdag", "Feestdag"],
        ["2026-12-26", "2e Kerstdag", "Feestdag"],
    ],
    columns=["Datum", "Omschrijving", "Type"],
)
st.dataframe(template_df, use_container_width=True, hide_index=True)
