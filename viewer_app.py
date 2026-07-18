
import os
from datetime import date
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
from shared import (
    APP_VERSION,
    MATERIAALTYPE_ORDER,
    MATERIAALTYPE_DAG_COLS,
    WITTE_VOORRAAD_STATUSSEN,
    OTIF_GAUGE_GREEN_THRESHOLD,
    OTIF_GAUGE_ORANGE_THRESHOLD,
    previous_workday,
    load_published_data,
    load_metadata,
    build_dashboard_data,
    compute_otif,
    get_peildatum_from_metadata,
    find_originele_leverdatum_column,
    format_int,
    format_pct,
    render_environment_banner,
    get_page_title,
    is_test_environment,
)

st.set_page_config(layout="wide", page_title=get_page_title("Capaciteitsplanning Coatinc Groningen"))

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


def make_materiaaltype_matplotlib_chart(day_df: pd.DataFrame):
    plot_df = day_df.copy()
    labels = plot_df["Label_nl"].tolist()
    x = np.arange(len(plot_df))
    capacity = plot_df["Capaciteit_kg"].tolist()
    is_holiday = plot_df["Is_feestdag_of_sluiting"].tolist()

    colors = {
        "Constructie": "#70AD47",
        "Maatwerk": "#2E75B6",
        "Seriewerk": "#8064A2",
        "Overig / onbekend": "#A6A6A6",
    }
    GRIJS = "#BDD7EE"

    fig, ax = plt.subplots(figsize=(11, 4.4))
    for i, holiday in enumerate(is_holiday):
        if holiday:
            ax.axvspan(i - 0.5, i + 0.5, alpha=0.12, color="red", zorder=0)

    # Capaciteitsbalk als achtergrond, gelijk aan de bestaande grafiek.
    ax.bar(x, capacity, width=0.56, color=GRIJS, alpha=0.6, label="Capaciteit", zorder=2)

    bottom = np.zeros(len(plot_df))
    segment_values = []
    for materiaaltype in MATERIAALTYPE_ORDER:
        col = MATERIAALTYPE_DAG_COLS[materiaaltype]
        values = plot_df[col].fillna(0).astype(float).to_numpy() if col in plot_df.columns else np.zeros(len(plot_df))
        segment_values.append(values)
        ax.bar(
            x,
            values,
            width=0.36,
            bottom=bottom,
            color=colors.get(materiaaltype, "#A6A6A6"),
            label=materiaaltype,
            zorder=3,
        )
        bottom = bottom + values

    load = bottom
    ax.set_title("Tonnage per materiaaltype op verzinkdatum", fontsize=15, pad=14)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=0)
    ax.set_ylabel("KG")
    ax.grid(axis="y", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_alpha(0.3)
    ax.spines["bottom"].set_alpha(0.3)
    ax.legend(frameon=False, ncols=5, loc="upper left")

    ymax = max(max(capacity) if capacity else 0, max(load) if len(load) else 0) * 1.15
    if ymax <= 0:
        ymax = 1
    ax.set_ylim(0, ymax)

    for i, val in enumerate(load):
        if val > 0:
            ax.text(i, val + ymax * 0.015, format_int(val), ha="center", va="bottom", fontsize=8, fontweight="bold")

    # Toon segmentlabels alleen als het segment groot genoeg is, zodat de grafiek leesbaar blijft.
    bottom_for_label = np.zeros(len(plot_df))
    for materiaaltype, values in zip(MATERIAALTYPE_ORDER, segment_values):
        for i, val in enumerate(values):
            if val > ymax * 0.07:
                ax.text(
                    i,
                    bottom_for_label[i] + val / 2,
                    format_int(val),
                    ha="center",
                    va="center",
                    fontsize=7,
                    color="white" if materiaaltype != "Overig / onbekend" else "black",
                )
        bottom_for_label = bottom_for_label + values

    fig.tight_layout()
    return fig


def make_otif_gauge(
    otif_pct: float | None,
    threshold_green: float = OTIF_GAUGE_GREEN_THRESHOLD,
    threshold_orange: float = OTIF_GAUGE_ORANGE_THRESHOLD,
    title: str = "OTIF",
    figsize: tuple[float, float] = (5.2, 3.4),
):
    """
    Teken een halfronde snelheidsmeter voor de OTIF-KPI.
    - < 80 %             → rood
    - 80 % – < 96 %      → oranje
    - ≥ 96 %             → groen (in lijn met de businessdoelstelling)
    """
    from matplotlib.patches import Wedge

    ROOD   = "#dc2626"
    ORANJE = "#f59e0b"
    GROEN  = "#16a34a"
    GRIJS  = "#e5e7eb"
    DONKER = "#111827"
    MIDGRIJS = "#6b7280"

    fig, ax = plt.subplots(figsize=figsize)
    ax.set_xlim(-1.25, 1.25)
    ax.set_ylim(-0.55, 1.25)
    ax.set_aspect("equal")
    ax.axis("off")

    # 0 % ligt links (180°), 100 % ligt rechts (0°).
    def pct_to_angle(p: float) -> float:
        return 180.0 - (max(0.0, min(100.0, p)) / 100.0) * 180.0

    outer_r = 1.0
    ring_width = 0.28

    # Achtergrondring als er nog geen waarde is.
    ax.add_patch(Wedge((0, 0), outer_r, 0, 180, width=ring_width, facecolor=GRIJS, edgecolor="none"))

    # Gekleurde zones.
    ax.add_patch(Wedge((0, 0), outer_r, pct_to_angle(threshold_orange), 180,
                       width=ring_width, facecolor=ROOD, edgecolor="none"))
    ax.add_patch(Wedge((0, 0), outer_r, pct_to_angle(threshold_green), pct_to_angle(threshold_orange),
                       width=ring_width, facecolor=ORANJE, edgecolor="none"))
    ax.add_patch(Wedge((0, 0), outer_r, 0, pct_to_angle(threshold_green),
                       width=ring_width, facecolor=GROEN, edgecolor="none"))

    # Schaalstreepjes en labels op 0, 20, 40, 60, 80, 100.
    for pct in (0, 20, 40, 60, 80, 100):
        angle_rad = np.deg2rad(pct_to_angle(pct))
        x1 = (outer_r - ring_width) * np.cos(angle_rad)
        y1 = (outer_r - ring_width) * np.sin(angle_rad)
        x2 = (outer_r - ring_width - 0.06) * np.cos(angle_rad)
        y2 = (outer_r - ring_width - 0.06) * np.sin(angle_rad)
        ax.plot([x1, x2], [y1, y2], color=DONKER, linewidth=1.1)
        xl = (outer_r - ring_width - 0.16) * np.cos(angle_rad)
        yl = (outer_r - ring_width - 0.16) * np.sin(angle_rad)
        ax.text(xl, yl, f"{pct}", ha="center", va="center", fontsize=8, color=DONKER)

    # Titel bovenin.
    ax.text(0, 1.14, title, ha="center", va="center", fontsize=13, fontweight="bold", color=DONKER)

    # Naaldwaarde + numerieke weergave.
    if otif_pct is None:
        ax.plot([0], [0], marker="o", color=MIDGRIJS, markersize=9)
        ax.text(0, -0.28, "geen data", ha="center", va="center",
                fontsize=14, fontweight="bold", color=MIDGRIJS)
        return fig

    clip = max(0.0, min(100.0, float(otif_pct)))
    needle_angle = np.deg2rad(pct_to_angle(clip))
    needle_len = outer_r - ring_width - 0.02
    nx = needle_len * np.cos(needle_angle)
    ny = needle_len * np.sin(needle_angle)
    ax.plot([0, nx], [0, ny], color=DONKER, linewidth=2.8, solid_capstyle="round")
    ax.plot([0], [0], marker="o", color=DONKER, markersize=11)

    # Kleur van het cijfer volgt de zonewaarin de waarde valt.
    if otif_pct >= threshold_green:
        num_color = GROEN
    elif otif_pct >= threshold_orange:
        num_color = ORANJE
    else:
        num_color = ROOD

    ax.text(0, -0.30, f"{otif_pct:.1f}%", ha="center", va="center",
            fontsize=24, fontweight="bold", color=num_color)

    fig.tight_layout()
    return fig


if os.path.exists("logo_coatinc_groningen.png"):
    st.sidebar.image("logo_coatinc_groningen.png", width=200)
st.sidebar.caption(APP_VERSION)
render_environment_banner("Viewer")

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

# ── OTIF-instellingen ────────────────────────────────────────────────────────
# De peildatum is standaard de dag van de laatste publicatie (het exportmoment
# van de brondata). De gebruiker kan hem zo nodig aanpassen voor terugkijkende
# analyses.
st.sidebar.markdown("---")
st.sidebar.subheader("OTIF")
default_peildatum = get_peildatum_from_metadata(meta)
otif_peildatum = st.sidebar.date_input(
    "Peildatum OTIF",
    value=default_peildatum,
    help=(
        "Standaard de dag waarop de brondata is geëxporteerd/gepubliceerd. "
        "Orders met deze datum in 'Datum' (of 'Originele datum') worden getoetst."
    ),
)

# Vergelijkingskolom kiezen. In de testomgeving mag tussen 'Datum' en
# 'Originele datum' worden geschakeld; in productie blijft de businessdefinitie
# 'Datum'. Als de kolom 'Originele_leverdatum' niet in de data zit, is de
# optie automatisch niet beschikbaar.
_originele_beschikbaar = (
    "Originele_leverdatum" in df.columns
    and pd.to_datetime(df["Originele_leverdatum"], errors="coerce").notna().any()
)

otif_datum_keuzes = {"Datum (leverdatum)": ("Leverdatum", "Datum")}
if _originele_beschikbaar:
    otif_datum_keuzes["Originele datum"] = ("Originele_leverdatum", "Originele datum")

if is_test_environment() and len(otif_datum_keuzes) > 1:
    otif_datum_keuze_label = st.sidebar.radio(
        "OTIF vergelijken met",
        list(otif_datum_keuzes.keys()),
        index=0,
        help=(
            "In de testomgeving kan zowel tegen de huidige 'Datum' als tegen de "
            "'Originele datum' getoetst worden om beide varianten te evalueren."
        ),
    )
else:
    # Buiten test of zonder Originele datum: default op de businessdefinitie.
    otif_datum_keuze_label = "Datum (leverdatum)"
    if is_test_environment() and not _originele_beschikbaar:
        st.sidebar.caption(
            "ℹ️ Originele datum niet aangetroffen in de weekexports; "
            "alleen vergelijking tegen 'Datum' beschikbaar."
        )

otif_date_column, otif_date_label = otif_datum_keuzes[otif_datum_keuze_label]

otif_result = compute_otif(
    df,
    peildatum=otif_peildatum,
    date_column=otif_date_column,
    date_column_label=otif_date_label,
)

tab1, tab2, tab3, tab4 = st.tabs(["Dashboard", "Gebruikte gegevens", "OTIF", "Debug"])

with tab1:
    st.subheader("Voorraadoverzicht en OTIF")

    # Zwarte voorraad: momentopname o.b.v. Status, los van de planningshorizon.
    # "Totale zwarte voorraad" = som van de twee subcategorieën hieronder.
    kg_productie_gereed = float(
        df.loc[df["Zwarte_voorraad_categorie"] == "Productie gereed", "Gewicht_effectief_kg"].sum()
    )
    kg_voorbewerking_geblokkeerd = float(
        df.loc[df["Zwarte_voorraad_categorie"] == "Binnengemeld/Voorbewerking/Geblokkeerd", "Gewicht_effectief_kg"].sum()
    )
    kg_totale_zwarte_voorraad = kg_productie_gereed + kg_voorbewerking_geblokkeerd

    # Witte voorraad: statussen die al afgehandeld/verzinkt zijn, maar nog als
    # voorraadstroom zichtbaar moeten zijn. Matching bewust case-insensitive.
    witte_statussen_norm = {str(s).strip().casefold() for s in WITTE_VOORRAAD_STATUSSEN}
    status_norm = df["Status"].astype(str).str.strip().str.casefold() if "Status" in df.columns else pd.Series([], dtype=str)
    kg_witte_voorraad = float(
        df.loc[status_norm.isin(witte_statussen_norm), "Gewicht_effectief_kg"].sum()
    ) if "Status" in df.columns else 0.0

    # Gedeelde opmaak voor de twee voorraadkaarten. De KPI-tegel voor OTIF
    # staat in een aparte Streamlit-kolom naast de kaarten.
    st.markdown(
        """
        <style>
        .cgr-stock-card {
            border: 1px solid #d0d7de;
            border-radius: 14px;
            padding: 1.05rem 1.15rem;
            background: #ffffff;
            box-shadow: 0 1px 2px rgba(16, 24, 40, 0.06);
            height: 100%;
            box-sizing: border-box;
        }
        .cgr-stock-card-accent-black { border-left: 8px solid #1f2937; }
        .cgr-stock-card-accent-white { border-left: 8px solid #94a3b8; }
        .cgr-stock-card-accent-otif  { border-left: 8px solid #16a34a; }
        .cgr-stock-label {
            font-size: 0.95rem;
            color: #475569;
            font-weight: 650;
            margin-bottom: 0.25rem;
        }
        .cgr-stock-value {
            font-size: 2.15rem;
            line-height: 1.12;
            color: #111827;
            font-weight: 750;
            margin-bottom: 0.7rem;
        }
        .cgr-stock-detail {
            font-size: 0.9rem;
            color: #475569;
            line-height: 1.45;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col_zwart, col_wit, col_otif = st.columns([1, 1, 1.25])

    with col_zwart:
        st.markdown(
            f"""
            <div class="cgr-stock-card cgr-stock-card-accent-black">
                <div class="cgr-stock-label">Zwarte voorraad (kg)</div>
                <div class="cgr-stock-value">{format_int(kg_totale_zwarte_voorraad)}</div>
                <div class="cgr-stock-detail">
                    Productie gereed: <strong>{format_int(kg_productie_gereed)}</strong> kg<br>
                    Binnengemeld / voorbewerking / geblokkeerd: <strong>{format_int(kg_voorbewerking_geblokkeerd)}</strong> kg
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_wit:
        st.markdown(
            f"""
            <div class="cgr-stock-card cgr-stock-card-accent-white">
                <div class="cgr-stock-label">Witte voorraad (kg)</div>
                <div class="cgr-stock-value">{format_int(kg_witte_voorraad)}</div>
                <div class="cgr-stock-detail">
                    Statussen: Afgehaald, Nabewerking nog uitvoeren, PC Afgehaald en Coat gereed
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_otif:
        # OTIF: snelheidsmeter op basis van de orders op de peildatum.
        peildatum_str = pd.Timestamp(otif_result["peildatum"]).strftime("%d-%m-%Y")
        gauge_title = f"OTIF · {otif_result['date_column_label']} = {peildatum_str}"
        st.pyplot(
            make_otif_gauge(otif_result["otif_pct"], title=gauge_title),
            clear_figure=True,
            use_container_width=True,
        )
        if otif_result["totaal"] > 0:
            samenvatting = (
                f"{otif_result['gereed']} op tijd · "
                f"{otif_result['niet_gereed']} te laat · "
                f"{otif_result['nvt']} nvt"
            )
            if otif_result["onbekend"] > 0:
                samenvatting += f" · {otif_result['onbekend']} onbekend"
            samenvatting += f" (totaal {otif_result['totaal']} orders)"
            st.caption(samenvatting)
        else:
            st.caption("Geen orders met deze peildatum in de dataset.")

    st.subheader("Eerstvolgende leverdatum")
    st.markdown(f'<div style="padding: 1rem 1.25rem; border-radius: 12px; border: 1px solid #d0d7de; background-color: #f6f8fa; margin-bottom: 0.75rem;"><div style="font-size: 2.2rem; font-weight: 700;">{advies_datum.strftime("%d-%m-%Y")}</div></div>', unsafe_allow_html=True)
    st.pyplot(make_professional_matplotlib_chart(dag), clear_figure=True, use_container_width=True)
    st.caption("De grafiek toont de geplande belasting per verzinkdatum, uitgesplitst naar bevestigde orders en reserveringen. De verzinkdatum is berekend als de leverdatum minus het ingestelde aantal werkdagen.")

    st.subheader("Materiaaltype per verzinkdatum")
    st.pyplot(make_materiaaltype_matplotlib_chart(dag), clear_figure=True, use_container_width=True)
    st.caption("Deze grafiek toont dezelfde totale dagbelasting, maar dan uitgesplitst naar materiaaltype op basis van het klantsegment uit debtor-export.xlsx.")

    st.subheader("Dagoverzicht")
    dag_display = dag.copy()
    dag_display["Verzinkdatum"] = dag_display["Verzinkdatum"].dt.date
    dag_segment_cols = [col for col in MATERIAALTYPE_DAG_COLS.values() if col in dag_display.columns]
    dag_display["Gewicht_kg"] = dag_display["Gewicht_kg"].apply(format_int)
    for c in dag_segment_cols:
        dag_display[c] = dag_display[c].apply(format_int)
    dag_display["Capaciteit_kg"] = dag_display["Capaciteit_kg"].apply(format_int)
    dag_display["Benutting_pct"] = dag_display["Benutting_pct"].apply(format_pct)
    dag_display["Traverses_berekend"] = dag_display["Traverses_berekend"].astype(int)
    dag_display["Dagtype"] = dag_display["Dagtype"].astype(str)
    def mark_holiday_row(row):
        if row["Dagtype"] == "Feestdag / sluiting":
            return ["background-color: rgba(220, 38, 38, 0.12)"] * len(row)
        return [""] * len(row)
    dag_cols = ["Verzinkdatum", "Dagtype", "Aantal_orders_te_verzinken", "Gewicht_kg"] + dag_segment_cols + ["Capaciteit_kg", "Benutting_pct", "Traverses_berekend", "Status"]
    st.dataframe(dag_display[dag_cols].style.apply(mark_holiday_row, axis=1), width="stretch", hide_index=True)
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
    relevant_cols = ["Bronbestand","Bron_week","Nummer","Ordernummer_base","Debiteurnummer","Klantnaam","Segment_debtor_export","Materiaaltype","Gewicht_effectief_kg","Status","Verzinkstatus","Meegeteld_in_planning","Reden_uitsluiting","Datum","Leverdatum","Verzinkdatum","Aanleveren depot","Gewicht","Gewicht_export_kg","Gewicht_order_kg","Regels_per_order","Gewicht_2g_verdeeld_kg","Gewicht_bron"]
    relevant_cols = [c for c in relevant_cols if c in df.columns]
    controle_df = df[relevant_cols].copy() if toon_alle_regels else df_plan[relevant_cols].copy()

    # Houd een interne datumkolom beschikbaar voor filtering, voordat we naar display-format omzetten.
    if "Verzinkdatum" in controle_df.columns:
        controle_df["_Verzinkdatum_filter"] = pd.to_datetime(controle_df["Verzinkdatum"], errors="coerce")

    st.markdown("### Filters")
    filter_col1, filter_col2, filter_col3, filter_col4, filter_col5 = st.columns([1.2, 1, 1, 1, 1.2])

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

    with filter_col5:
        if "Materiaaltype" in controle_df.columns:
            materiaaltype_opties = [m for m in MATERIAALTYPE_ORDER if m in set(controle_df["Materiaaltype"].dropna().astype(str))]
            overige_opties = sorted([m for m in controle_df["Materiaaltype"].dropna().astype(str).unique() if m not in materiaaltype_opties])
            materiaaltype_opties = materiaaltype_opties + overige_opties
            selected_materiaaltype = st.multiselect(
                "Materiaaltype",
                options=materiaaltype_opties,
                default=materiaaltype_opties,
            )
        else:
            selected_materiaaltype = []

    filtered_df = controle_df.copy()

    if "Status" in filtered_df.columns and selected_status:
        filtered_df = filtered_df[filtered_df["Status"].astype(str).isin(selected_status)]

    if "Verzinkstatus" in filtered_df.columns and selected_verzinkstatus:
        filtered_df = filtered_df[filtered_df["Verzinkstatus"].astype(str).isin(selected_verzinkstatus)]

    if "Materiaaltype" in filtered_df.columns and selected_materiaaltype:
        filtered_df = filtered_df[filtered_df["Materiaaltype"].astype(str).isin(selected_materiaaltype)]

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
    st.subheader("OTIF – On Time In Full")

    peildatum_str = pd.Timestamp(otif_result["peildatum"]).strftime("%d-%m-%Y")
    st.caption(
        f"Peildatum: **{peildatum_str}** · vergeleken met **{otif_result['date_column_label']}** · "
        f"reserveringen worden niet meegeteld."
    )

    # Kerngetallen boven de tabel.
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
    kpi_col1.metric("Totaal orders", f"{otif_result['totaal']:,}".replace(",", "."))
    kpi_col2.metric("Gereed (Ja)", f"{otif_result['gereed']:,}".replace(",", "."))
    kpi_col3.metric("Te laat (Nee)", f"{otif_result['niet_gereed']:,}".replace(",", "."))
    kpi_col4.metric("Nvt", f"{otif_result['nvt']:,}".replace(",", "."))
    otif_pct_str = "n.v.t." if otif_result["otif_pct"] is None else f"{otif_result['otif_pct']:.1f}%"
    kpi_col5.metric("OTIF", otif_pct_str)

    # Gauge groter tonen naast een korte formule-uitleg.
    gauge_col, uitleg_col = st.columns([1, 1.1])
    with gauge_col:
        st.pyplot(
            make_otif_gauge(
                otif_result["otif_pct"],
                title=f"OTIF · {otif_result['date_column_label']} = {peildatum_str}",
                figsize=(6.0, 4.0),
            ),
            clear_figure=True,
            use_container_width=True,
        )
    with uitleg_col:
        st.markdown(
            f"""
**Berekening**  
OTIF = (1 − *aantal te laat* / *totaal aantal orders*) × 100 %  
= (1 − {otif_result['niet_gereed']} / {max(otif_result['totaal'], 1)}) × 100 %  
= **{otif_pct_str}**

**Status → OTIF-mapping** (zie Status_OTIF.xlsx)
- **Ja** (op tijd): uitgeleverd, Afgehaald, Coat gereed, UB V Gereed
- **Nee** (te laat): Productie gereed, Geblokkeerd, Opgehangen, PC Opgehangen, UB, Nabewerking nog uitvoeren, meetrapport
- **Nvt**: PC Afgehaald

Groen op de meter vanaf **{OTIF_GAUGE_GREEN_THRESHOLD:.0f}%**, oranje vanaf **{OTIF_GAUGE_ORANGE_THRESHOLD:.0f}%**, onder deze grens rood.
"""
        )

    st.markdown("---")
    st.subheader("Gebruikte gegevens (OTIF)")

    otif_orders = otif_result["orders"]
    if otif_orders.empty:
        st.info(
            "Geen orders gevonden met de gekozen peildatum in "
            f"'{otif_result['date_column_label']}'. Controleer eventueel of de "
            "juiste peildatum in de zijbalk is geselecteerd."
        )
    else:
        # Kolomkeuze: aansluiten bij 'Gebruikte gegevens' voor herkenbaarheid.
        otif_cols = [
            "Bronbestand", "Bron_week", "Nummer", "Ordernummer_base",
            "Debiteurnummer", "Klantnaam",
            "Segment_debtor_export", "Materiaaltype",
            "Datum", "Leverdatum", "Originele_leverdatum",
            "Status", "OTIF_status", "Verzinkstatus",
            "Gewicht_effectief_kg", "Gewicht_bron",
        ]
        otif_cols = [c for c in otif_cols if c in otif_orders.columns]
        otif_orders_display = otif_orders[otif_cols].copy()

        # Filter op OTIF-status voor snelle drill-down.
        f_col1, f_col2 = st.columns([1, 3])
        with f_col1:
            otif_status_opties = ["Ja", "Nee", "Nvt", "Onbekend"]
            beschikbare_status = [s for s in otif_status_opties if s in set(otif_orders_display["OTIF_status"].astype(str))]
            selected_otif_status = st.multiselect(
                "OTIF-status",
                options=beschikbare_status,
                default=beschikbare_status,
            )

        if selected_otif_status:
            otif_orders_display = otif_orders_display[otif_orders_display["OTIF_status"].astype(str).isin(selected_otif_status)]

        # Datums netjes weergeven.
        for c in ("Datum", "Leverdatum", "Originele_leverdatum"):
            if c in otif_orders_display.columns:
                otif_orders_display[c] = pd.to_datetime(otif_orders_display[c], errors="coerce").dt.date
        if "Gewicht_effectief_kg" in otif_orders_display.columns:
            otif_orders_display["Gewicht_effectief_kg"] = otif_orders_display["Gewicht_effectief_kg"].round(2)

        def _color_otif_row(row):
            status_value = str(row.get("OTIF_status", ""))
            if status_value == "Ja":
                return ["background-color: rgba(22, 163, 74, 0.10)"] * len(row)
            if status_value == "Nee":
                return ["background-color: rgba(220, 38, 38, 0.12)"] * len(row)
            if status_value == "Nvt":
                return ["background-color: rgba(148, 163, 184, 0.15)"] * len(row)
            return [""] * len(row)

        st.caption(f"Getoonde regels: {len(otif_orders_display):,}".replace(",", "."))
        st.dataframe(
            otif_orders_display.style.apply(_color_otif_row, axis=1),
            width="stretch",
            hide_index=True,
        )

        # ── Excel-download van de OTIF-tabel ──────────────────────────────
        import io
        otif_pct_excel = (
            "n.v.t." if otif_result["otif_pct"] is None
            else round(otif_result["otif_pct"], 2)
        )
        otif_xlsx_buffer = io.BytesIO()
        with pd.ExcelWriter(otif_xlsx_buffer, engine="openpyxl") as writer:
            otif_orders_display.to_excel(writer, index=False, sheet_name="OTIF-orders")
            # Samenvatting op tweede tabblad, zodat het bestand op zichzelf leesbaar is.
            samenvatting_df = pd.DataFrame(
                [
                    {"Metric": "Peildatum",             "Waarde": peildatum_str},
                    {"Metric": "Vergeleken met",        "Waarde": otif_result["date_column_label"]},
                    {"Metric": "Totaal orders",         "Waarde": otif_result["totaal"]},
                    {"Metric": "Gereed (Ja)",           "Waarde": otif_result["gereed"]},
                    {"Metric": "Te laat (Nee)",         "Waarde": otif_result["niet_gereed"]},
                    {"Metric": "Nvt",                   "Waarde": otif_result["nvt"]},
                    {"Metric": "Onbekend",              "Waarde": otif_result["onbekend"]},
                    {"Metric": "OTIF (%)",              "Waarde": otif_pct_excel},
                    {"Metric": "Drempel groen (%)",     "Waarde": OTIF_GAUGE_GREEN_THRESHOLD},
                    {"Metric": "Drempel oranje (%)",    "Waarde": OTIF_GAUGE_ORANGE_THRESHOLD},
                ]
            )
            samenvatting_df.to_excel(writer, index=False, sheet_name="Samenvatting")
        otif_xlsx_buffer.seek(0)
        st.download_button(
            label="📥 Download OTIF-gegevens als Excel (.xlsx)",
            data=otif_xlsx_buffer,
            file_name=(
                f"otif_{otif_result['date_column']}_"
                f"{pd.Timestamp(otif_result['peildatum']).strftime('%Y%m%d')}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab4:
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
    if "Materiaaltype" in df_plan.columns:
        for materiaaltype in MATERIAALTYPE_ORDER:
            debug_rows.append({
                "Categorie":"Gewicht per materiaaltype",
                "Omschrijving":materiaaltype,
                "Waarde":int(round(df_plan.loc[df_plan["Materiaaltype"] == materiaaltype, "Gewicht_effectief_kg"].sum(), 0)),
            })
    for _, row in export_file_summary.iterrows():
        debug_rows.append({"Categorie":"Bestanden","Omschrijving":f'{row["Bronbestand"]} ({row["Bron_week"]})',"Waarde":int(row["Aantal_regels_ingelezen"])})
    debug_df = pd.DataFrame(debug_rows)
    debug_df["Waarde"] = debug_df["Waarde"].astype(str)
    st.dataframe(debug_df, width="stretch", hide_index=True)
