import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st

from shared import (
    APP_VERSION,
    REQUIRED_FILES,
    load_metadata,
    publish_files,
    save_metadata,
    upload_file,
    validate_feestdagen_xlsx,
    validate_required_files_in_folder,
)
from pathlib import Path
import tempfile

st.set_page_config(layout="wide", page_title="Capaciteitsplanning Coatinc Groningen - Beheer")

if os.path.exists("logo_coatinc_groningen.png"):
    st.sidebar.image("logo_coatinc_groningen.png", use_container_width=True)
st.sidebar.caption(APP_VERSION)

password = st.sidebar.text_input("Wachtwoord", type="password")
if password != "coatinc2026":
    st.stop()

st.title("Capaciteitsplanning Coatinc Groningen – Beheer")
st.caption("Upload per bestand en publiceer daarna de volledige dataset voor alle kijkers.")

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
            "Nieuw geüpload": "✅ Ja" if uploads.get(fname) is not None else "–",
            "Staat al in cloud": "✅ Ja" if fname in cloud_files else "❌ Nee",
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
