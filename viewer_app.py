
import os
from datetime import date
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
from shared import APP_VERSION, previous_workday, load_published_data, load_metadata, build_dashboard_data, format_int, format_pct

st.set_page_config(layout="wide", page_title="Capaciteitsplanning Coatinc Groningen")

def make_professional_matplotlib_chart(day_df: pd.DataFrame):
    plot_df = day_df.copy()
    labels = plot_df["Label_nl"].tolist()
    x = np.arange(len(plot_df))
    capacity   = plot_df["Capaciteit_kg"].tolist()
    definitief = plot_df["Gewicht_definitief_kg"].tolist()
    reservering = plot_df["Gewicht_reservering_kg"].tolist()
    load       = plot_df["Gewicht_kg"].tolist()
    is_holiday = plot_df["Is_feestdag_of_sluiting"].tolist()

    BLAUW      = "#2E75B6"
    ORANJE     = "#F4A623"
    GRIJS      = "#BDD7EE"

    fig, ax = plt.subplots(figsize=(11, 4.8))
    for i, holiday in enumerate(is_holiday):
        if holiday:
            ax.axvspan(i - 0.5, i + 0.5, alpha=0.12, color="red", zorder=0)

    # Capaciteitsbalk (achtergrond)
    ax.bar(x, capacity, width=0.56, color=GRIJS, alpha=0.6, label="Capaciteit", zorder=2)

    # Gestapelde dagbelasting: definitief + reservering
    bars_def = ax.bar(x, definitief, width=0.36, color=BLAUW, label="Bevestigde orders", zorder=3)
    bars_res = ax.bar(x, reservering, width=0.36, bottom=definitief, color=ORANJE, label="Reserveringen", zorder=3)

    ax.set_title("Capaciteit versus dagbelasting op verzinkdatum", fontsize=15, pad=14)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=0)
    ax.set_ylabel("KG")
    ax.grid(axis="y", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_alpha(0.3)
    ax.spines["bottom"].set_alpha(0.3)
    ax.legend(frameon=False, ncols=3, loc="upper left")

    ymax = max(max(capacity) if capacity else 0, max(load) if load else 0) * 1.15
    if ymax <= 0:
        ymax = 1
    ax.set_ylim(0, ymax)

    # Totaallabel boven de volledige balk
    for i, val in enumerate(load):
        if val > 0:
            ax.text(i, val + ymax * 0.015, format_int(val), ha="center", va="bottom", fontsize=8, fontweight="bold")

    # Klein label per type binnen de balk
    for i, (d, r) in enumerate(zip(definitief, reservering)):
        if d > ymax * 0.06:  # alleen tonen als segment groot genoeg is
            ax.text(i, d / 2, format_int(d), ha="center", va="center", fontsize=7, color="white")
        if r > ymax * 0.06:
            ax.text(i, d + r / 2, format_int(r), ha="center", va="center", fontsize=7, color="white")
    fig.tight_layout()
    return fig

if os.path.exists("logo_coatinc_groningen.png"):
    st.sidebar.image("logo_coatinc_groningen.png", width=200)
st.sidebar.caption(APP_VERSION)

meta = load_metadata()
if meta:
    st.sidebar.markdown("**Laatste publicatie**")
    last_update = meta.get("published_at", "-")
    st.sidebar.write(f"Laatste update: {last_update}")
    published_by = meta.get("published_by", "")
    notes = meta.get("notes", "")
    if published_by:
        st.sidebar.write(f"Door: {published_by}")
    if notes:
        st.sidebar.write(f"Toelichting: {notes}")

st.title("Capaciteitsplanning Coatinc Groningen")
st.sidebar.header("Instellingen")
capaciteit_ton = st.sidebar.slider("Max capaciteit per dag (ton)", 50, 90, 60, 5)
capaciteit_kg = capaciteit_ton * 1000
offset = st.sidebar.selectbox("Verzinkdatum = leverdatum - X werkdagen", [1, 2, 3, 4], index=1)
kg_per_traverse = st.sidebar.number_input("KG per traverse", min_value=100, max_value=10000, value=1000, step=100)
default_start = previous_workday(date.today())
startdatum = st.sidebar.date_input("Startdatum rapport", value=default_start)
toon_alle_regels = st.sidebar.checkbox("Toon alle regels in controletab", value=True)

try:
    df_raw, export_file_summary, order_file, holiday_df = load_published_data()
except Exception as e:
    st.error(f"Kan gepubliceerde data niet laden: {e}")
    st.stop()

df, df_plan, dag, week, advies_datum = build_dashboard_data(df_raw, holiday_df, startdatum, capaciteit_kg, offset, kg_per_traverse)
tab1, tab2, tab3 = st.tabs(["Dashboard", "Gebruikte gegevens", "Debug"])

with tab1:
    st.subheader("KPI-overzicht")
    k1, k2 = st.columns(2)
    aantal_orders_te_verzinken = int((df["Meegeteld_in_planning"] == "Ja").sum())
    totaal_kg_te_verzinken = float(df_plan["Gewicht_effectief_kg"].sum()) if len(df_plan) > 0 else 0.0
    with k1:
        st.metric("Aantal orders te verzinken", f"{aantal_orders_te_verzinken:,}".replace(",", "."))
    with k2:
        st.metric("Totaal KG te verzinken", f"{int(round(totaal_kg_te_verzinken, 0)):,}".replace(",", "."))

    # Zwarte voorraad: momentopname o.b.v. Status, los van de planningshorizon.
    # "Totale zwarte voorraad" = som van de twee subcategorieën hieronder.
    kg_productie_gereed = float(
        df.loc[df["Zwarte_voorraad_categorie"] == "Productie gereed", "Gewicht_effectief_kg"].sum()
    )
    kg_voorbewerking_geblokkeerd = float(
        df.loc[df["Zwarte_voorraad_categorie"] == "Binnengemeld/Voorbewerking/Geblokkeerd", "Gewicht_effectief_kg"].sum()
    )
    kg_totale_zwarte_voorraad = kg_productie_gereed + kg_voorbewerking_geblokkeerd

    zv1, zv2, zv3 = st.columns(3)
    with zv1:
        st.metric("Totale zwarte voorraad (kg)", format_int(kg_totale_zwarte_voorraad))
    with zv2:
        st.metric("Waarvan Productie gereed (kg)", format_int(kg_productie_gereed))
    with zv3:
        st.metric("Waarvan Binnengemeld/Voorbewerking/Geblokkeerd (kg)", format_int(kg_voorbewerking_geblokkeerd))

    st.subheader("Eerstvolgende leverdatum")
    st.markdown(f'<div style="padding: 1rem 1.25rem; border-radius: 12px; border: 1px solid #d0d7de; background-color: #f6f8fa; margin-bottom: 0.75rem;"><div style="font-size: 2.2rem; font-weight: 700;">{advies_datum.strftime("%d-%m-%Y")}</div></div>', unsafe_allow_html=True)
    st.pyplot(make_professional_matplotlib_chart(dag), clear_figure=True, use_container_width=True)
    st.caption("De grafiek toont de geplande belasting per verzinkdatum. De verzinkdatum is berekend als de leverdatum minus het ingestelde aantal werkdagen.")
    st.subheader("Dagoverzicht")
    dag_display = dag.copy()
    dag_display["Verzinkdatum"] = dag_display["Verzinkdatum"].dt.date
    dag_display["Gewicht_kg"] = dag_display["Gewicht_kg"].apply(format_int)
    dag_display["Capaciteit_kg"] = dag_display["Capaciteit_kg"].apply(format_int)
    dag_display["Benutting_pct"] = dag_display["Benutting_pct"].apply(format_pct)
    dag_display["Traverses_berekend"] = dag_display["Traverses_berekend"].astype(int)
    dag_display["Dagtype"] = dag_display["Dagtype"].astype(str)
    def mark_holiday_row(row):
        if row["Dagtype"] == "Feestdag / sluiting":
            return ["background-color: rgba(220, 38, 38, 0.12)"] * len(row)
        return [""] * len(row)
    st.dataframe(dag_display[["Verzinkdatum","Dagtype","Aantal_orders_te_verzinken","Gewicht_kg","Capaciteit_kg","Benutting_pct","Traverses_berekend","Status"]].style.apply(mark_holiday_row, axis=1), width="stretch", hide_index=True)
    st.subheader("Weekoverzicht")
    week_display = week.copy()
    week_display["Gewicht_kg"] = week_display["Gewicht_kg"].apply(format_int)
    week_display["Capaciteit_kg"] = week_display["Capaciteit_kg"].apply(format_int)
    week_display["Benutting_pct"] = week_display["Benutting_pct"].apply(format_pct)
    week_display["Traverses_berekend"] = week_display["Traverses_berekend"].astype(int)
    st.dataframe(week_display[["Jaar","Week","Aantal_orders_te_verzinken","Gewicht_kg","Capaciteit_kg","Benutting_pct","Traverses_berekend","Status"]], width="stretch", hide_index=True)
    st.subheader("Feestdagen en fabriekssluiting")
    holiday_show = holiday_df.copy()
    holiday_show["Datum"] = pd.to_datetime(holiday_show["Datum"]).dt.strftime("%d-%m-%Y")
    st.dataframe(holiday_show, width="stretch", hide_index=True)

with tab2:
    st.subheader("Gebruikte gegevens / controletabel")
    relevant_cols = ["Bronbestand","Bron_week","Nummer","Ordernummer_base","Klantnaam","Gewicht_effectief_kg","Status","Verzinkstatus","Meegeteld_in_planning","Reden_uitsluiting","Datum","Leverdatum","Verzinkdatum","Aanleveren depot","Gewicht","Gewicht_export_kg","Gewicht_order_kg","Regels_per_order","Gewicht_2g_verdeeld_kg","Gewicht_bron"]
    relevant_cols = [c for c in relevant_cols if c in df.columns]
    controle_df = df[relevant_cols].copy() if toon_alle_regels else df_plan[relevant_cols].copy()

    # Houd een interne datumkolom beschikbaar voor filtering, voordat we naar display-format omzetten.
    if "Verzinkdatum" in controle_df.columns:
        controle_df["_Verzinkdatum_filter"] = pd.to_datetime(controle_df["Verzinkdatum"], errors="coerce")

    st.markdown("### Filters")
    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns([1.2, 1, 1, 1])

    with filter_col1:
        if "Status" in controle_df.columns:
            status_options = sorted([str(x) for x in controle_df["Status"].dropna().unique()])
            selected_status = st.multiselect(
                "Status",
                options=status_options,
                default=status_options,
                help="Selecteer één of meerdere statussen. Alleen regels met deze status worden getoond.",
            )
        else:
            selected_status = []

    with filter_col2:
        if "_Verzinkdatum_filter" in controle_df.columns and controle_df["_Verzinkdatum_filter"].notna().any():
            min_verzinkdatum = controle_df["_Verzinkdatum_filter"].min().date()
            max_verzinkdatum = controle_df["_Verzinkdatum_filter"].max().date()
            verzinkdatum_van = st.date_input("Verzinkdatum vanaf", value=min_verzinkdatum)
        else:
            verzinkdatum_van = None

    with filter_col3:
        if "_Verzinkdatum_filter" in controle_df.columns and controle_df["_Verzinkdatum_filter"].notna().any():
            verzinkdatum_tot = st.date_input("Verzinkdatum t/m", value=max_verzinkdatum)
        else:
            verzinkdatum_tot = None

    with filter_col4:
        if "Verzinkstatus" in controle_df.columns:
            verzinkstatus_opties = sorted([str(x) for x in controle_df["Verzinkstatus"].dropna().unique()])
            selected_verzinkstatus = st.multiselect(
                "Verzinkstatus",
                options=verzinkstatus_opties,
                default=verzinkstatus_opties,
            )
        else:
            selected_verzinkstatus = []

    filtered_df = controle_df.copy()

    if "Status" in filtered_df.columns and selected_status:
        filtered_df = filtered_df[filtered_df["Status"].astype(str).isin(selected_status)]

    if "Verzinkstatus" in filtered_df.columns and selected_verzinkstatus:
        filtered_df = filtered_df[filtered_df["Verzinkstatus"].astype(str).isin(selected_verzinkstatus)]

    if "_Verzinkdatum_filter" in filtered_df.columns:
        if verzinkdatum_van is not None:
            filtered_df = filtered_df[filtered_df["_Verzinkdatum_filter"].dt.date >= verzinkdatum_van]
        if verzinkdatum_tot is not None:
            filtered_df = filtered_df[filtered_df["_Verzinkdatum_filter"].dt.date <= verzinkdatum_tot]

    st.caption(f"Getoonde regels: {len(filtered_df):,}".replace(",", "."))

    display_df = filtered_df.copy()
    if "Datum" in display_df.columns:
        display_df["Datum"] = pd.to_datetime(display_df["Datum"], errors="coerce").dt.date
    if "Leverdatum" in display_df.columns:
        display_df["Leverdatum"] = pd.to_datetime(display_df["Leverdatum"], errors="coerce").dt.date
    if "Verzinkdatum" in display_df.columns:
        display_df["Verzinkdatum"] = pd.to_datetime(display_df["Verzinkdatum"], errors="coerce").dt.date
    for c in ["Gewicht_export_kg","Gewicht_order_kg","Gewicht_2g_verdeeld_kg","Gewicht_effectief_kg"]:
        if c in display_df.columns:
            display_df[c] = display_df[c].round(2)

    if "_Verzinkdatum_filter" in display_df.columns:
        display_df = display_df.drop(columns=["_Verzinkdatum_filter"])

    st.dataframe(display_df, width="stretch", hide_index=True)

    # ── xlsx-export ──────────────────────────────────────────────────────
    # Export gebruikt dezelfde filters als de tabel op het scherm.
    import io
    xlsx_buffer = io.BytesIO()
    with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
        display_df.to_excel(writer, index=False, sheet_name="Gebruikte gegevens")
    xlsx_buffer.seek(0)
    st.download_button(
        label="📥 Download gefilterde tabel als Excel (.xlsx)",
        data=xlsx_buffer,
        file_name=f"capaciteitsplanning_export_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab3:
    st.subheader("Debug samenvatting")
    debug_rows = [
        {"Categorie":"Instellingen","Omschrijving":"Startdatum rapport","Waarde":str(startdatum)},
        {"Categorie":"Instellingen","Omschrijving":"Capaciteit per dag (kg)","Waarde":int(capaciteit_kg)},
        {"Categorie":"Instellingen","Omschrijving":"KG per traverse","Waarde":int(kg_per_traverse)},
        {"Categorie":"Bestanden","Omschrijving":"Orderbestand","Waarde":order_file},
        {"Categorie":"Kalender","Omschrijving":"Aantal feestdagen / sluitingen","Waarde":int(len(holiday_df))},
        {"Categorie":"Records","Omschrijving":"Niet verzinkt","Waarde":int((df["Verzinkstatus"] == "Niet verzinkt").sum())},
        {"Categorie":"Records","Omschrijving":"Voor startdatum rapport","Waarde":int(((df["Verzinkstatus"] == "Niet verzinkt") & (df["Verzinkdatum"] < pd.Timestamp(startdatum).normalize())).sum())},
        {"Categorie":"Records","Omschrijving":"Open orders","Waarde":int((df["Meegeteld_in_planning"] == "Ja").sum())},
        {"Categorie":"Gewicht","Omschrijving":"Totaal gewicht open orders (kg)","Waarde":int(round(df_plan["Gewicht_effectief_kg"].sum(), 0)) if len(df_plan) > 0 else 0},
    ]
    for _, row in export_file_summary.iterrows():
        debug_rows.append({"Categorie":"Bestanden","Omschrijving":f'{row["Bronbestand"]} ({row["Bron_week"]})',"Waarde":int(row["Aantal_regels_ingelezen"])})
    debug_df = pd.DataFrame(debug_rows)
    debug_df["Waarde"] = debug_df["Waarde"].astype(str)
    st.dataframe(debug_df, width="stretch", hide_index=True)
