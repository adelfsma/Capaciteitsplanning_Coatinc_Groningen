
import os
import shutil
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st
from shared import APP_VERSION, REQUIRED_FILES, PUBLISHED_DIR, ensure_dirs, validate_feestdagen_xlsx, validate_required_files_in_folder, save_metadata, load_metadata

st.set_page_config(layout="wide", page_title="Capaciteitsplanning Coatinc Groningen - Beheer")
if os.path.exists("logo_coatinc_groningen.png"):
    st.sidebar.image("logo_coatinc_groningen.png", use_container_width=True)
st.sidebar.caption(APP_VERSION)

password = st.sidebar.text_input("Wachtwoord", type="password")
if password != "coatinc2026":
    st.stop()
st.title("Capaciteitsplanning Coatinc Groningen – Beheer")
st.caption("Upload per bestand en publiceer daarna de volledige dataset voor alle kijkers.")
ensure_dirs()
meta = load_metadata()
top1, top2 = st.columns([1,2])
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
st.subheader("1. Upload bestanden")
st.caption("Upload elk bestand in het juiste vak. Publiceren vervangt alleen de bestanden die je hebt geüpload in deze sessie.")
uploads = {}
upload_cols = st.columns(2)
for i, fname in enumerate(REQUIRED_FILES):
    with upload_cols[i % 2]:
        uploads[fname] = st.file_uploader(f"Upload {fname}", type=["xlsx"], key=fname)

st.markdown("---")
st.subheader("2. Validatie")
validation_rows = []
for fname in REQUIRED_FILES:
    file_obj = uploads.get(fname)
    validation_rows.append({"Bestand": fname, "Nieuw geüpload": "Ja" if file_obj is not None else "Nee", "Bestaat al gepubliceerd": "Ja" if (PUBLISHED_DIR / fname).exists() else "Nee"})
st.dataframe(pd.DataFrame(validation_rows), use_container_width=True, hide_index=True)

publisher = st.text_input("Naam beheerder", value="")
notes = st.text_input("Notitie / omschrijving update", value="")
publish = st.button("Publiceer dataset", type="primary", use_container_width=True)

if publish:
    stage_dir = Path("staging_upload")
    if stage_dir.exists():
        shutil.rmtree(stage_dir)
    stage_dir.mkdir(parents=True, exist_ok=True)
    for fname in REQUIRED_FILES:
        existing = PUBLISHED_DIR / fname
        if existing.exists():
            shutil.copyfile(existing, stage_dir / fname)
    for fname, uploaded in uploads.items():
        if uploaded is not None:
            with open(stage_dir / fname, "wb") as f:
                f.write(uploaded.getbuffer())
    try:
        validate_required_files_in_folder(stage_dir)
        validate_feestdagen_xlsx(stage_dir / "feestdagen.xlsx")
        for fname in REQUIRED_FILES:
            shutil.copyfile(stage_dir / fname, PUBLISHED_DIR / fname)
        metadata = {
            "published_at": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
            "published_by": publisher if publisher else "Onbekend",
            "notes": notes if notes else "",
            "files": REQUIRED_FILES,
        }
        save_metadata(metadata)
        st.success("Dataset succesvol gepubliceerd.")
    except Exception as e:
        st.error(f"Publicatie mislukt: {e}")
    finally:
        if stage_dir.exists():
            shutil.rmtree(stage_dir)

st.markdown("---")
st.subheader("3. Template feestdagen.xlsx")
template_df = pd.DataFrame([
    ["2026-01-01","Nieuwjaarsdag","Feestdag"],
    ["2026-04-03","Goede vrijdag","Feestdag"],
    ["2026-04-06","2e Paasdag","Feestdag"],
    ["2026-04-27","Koningsdag","Feestdag"],
    ["2026-05-05","Bevrijdingsdag","Feestdag"],
    ["2026-05-14","Hemelvaartsdag","Feestdag"],
    ["2026-05-25","2e Pinksterdag","Feestdag"],
    ["2026-12-25","1e Kerstdag","Feestdag"],
    ["2026-12-26","2e Kerstdag","Feestdag"],
], columns=["Datum","Omschrijving","Type"])
st.dataframe(template_df, use_container_width=True, hide_index=True)
