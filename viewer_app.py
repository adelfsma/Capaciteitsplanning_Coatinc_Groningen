
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
    capacity = plot_df["Capaciteit_kg"].tolist()
    load = plot_df["Gewicht_kg"].tolist()
    is_holiday = plot_df["Is_feestdag_of_sluiting"].tolist()
    fig, ax = plt.subplots(figsize=(11, 4.8))
    for i, holiday in enumerate(is_holiday):
        if holiday:
            ax.axvspan(i - 0.5, i + 0.5, alpha=0.12, color="red", zorder=0)
    ax.bar(x, capacity, width=0.56, alpha=0.35, label="Capaciteit", zorder=2)
    ax.bar(x, load, width=0.36, label="Dagbelasting", zorder=3)
    ax.set_title("Capaciteit versus dagbelasting", fontsize=15, pad=14)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=0)
    ax.set_ylabel("KG")
    ax.grid(axis="y", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_alpha(0.3)
    ax.spines["bottom"].set_alpha(0.3)
    ax.legend(frameon=False, ncols=2, loc="upper left")
    ymax = max(max(capacity) if capacity else 0, max(load) if load else 0) * 1.15
    if ymax <= 0:
        ymax = 1
    ax.set_ylim(0, ymax)
    for i, val in enumerate(load):
        if val > 0:
            ax.text(i, val + ymax * 0.015, format_int(val), ha="center", va="bottom", fontsize=8)
    fig.tight_layout()
    return fig

if os.path.exists("logo_coatinc_groningen.png"):
    st.sidebar.image("logo_coatinc_groningen.png", use_container_width=True)
st.sidebar.caption(APP_VERSION)

password = st.sidebar.text_input("Wachtwoord", type="password")
if password != "coatinc2026":
    st.stop()
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
capaciteit_ton = st.sidebar.slider("Max capaciteit per dag (ton)", 50, 90, 70, 5)
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
    st.subheader("Eerstvolgende leverdatum")
    st.markdown(f'<div style="padding: 1rem 1.25rem; border-radius: 12px; border: 1px solid #d0d7de; background-color: #f6f8fa; margin-bottom: 0.75rem;"><div style="font-size: 2.2rem; font-weight: 700;">{advies_datum.strftime("%d-%m-%Y")}</div></div>', unsafe_allow_html=True)
    st.pyplot(make_professional_matplotlib_chart(dag), clear_figure=True, use_container_width=True)
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
    st.dataframe(dag_display[["Verzinkdatum","Dagtype","Aantal_orders_te_verzinken","Gewicht_kg","Capaciteit_kg","Benutting_pct","Traverses_berekend","Status"]].style.apply(mark_holiday_row, axis=1), use_container_width=True, hide_index=True)
    st.subheader("Weekoverzicht")
    week_display = week.copy()
    week_display["Gewicht_kg"] = week_display["Gewicht_kg"].apply(format_int)
    week_display["Capaciteit_kg"] = week_display["Capaciteit_kg"].apply(format_int)
    week_display["Benutting_pct"] = week_display["Benutting_pct"].apply(format_pct)
    week_display["Traverses_berekend"] = week_display["Traverses_berekend"].astype(int)
    st.dataframe(week_display[["Jaar","Week","Aantal_orders_te_verzinken","Gewicht_kg","Capaciteit_kg","Benutting_pct","Traverses_berekend","Status"]], use_container_width=True, hide_index=True)
    st.subheader("Feestdagen en fabriekssluiting")
    holiday_show = holiday_df.copy()
    holiday_show["Datum"] = pd.to_datetime(holiday_show["Datum"]).dt.strftime("%d-%m-%Y")
    st.dataframe(holiday_show, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Gebruikte gegevens / controletabel")
    relevant_cols = ["Bronbestand","Bron_week","Nummer","Ordernummer_base","Status","Verzinkstatus","Meegeteld_in_planning","Reden_uitsluiting","Datum","Leverdatum","Verzinkdatum","Gewicht","Gewicht_export_kg","Gewicht_order_kg","Regels_per_order","Gewicht_2g_verdeeld_kg","Gewicht_bron","Gewicht_effectief_kg"]
    relevant_cols = [c for c in relevant_cols if c in df.columns]
    controle_df = df[relevant_cols].copy() if toon_alle_regels else df_plan[relevant_cols].copy()
    if "Leverdatum" in controle_df.columns:
        controle_df["Leverdatum"] = pd.to_datetime(controle_df["Leverdatum"], errors="coerce").dt.date
    if "Verzinkdatum" in controle_df.columns:
        controle_df["Verzinkdatum"] = pd.to_datetime(controle_df["Verzinkdatum"], errors="coerce").dt.date
    for c in ["Gewicht_export_kg","Gewicht_order_kg","Gewicht_2g_verdeeld_kg","Gewicht_effectief_kg"]:
        if c in controle_df.columns:
            controle_df[c] = controle_df[c].round(2)
    st.dataframe(controle_df, use_container_width=True, hide_index=True)

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
    st.dataframe(pd.DataFrame(debug_rows), use_container_width=True, hide_index=True)
