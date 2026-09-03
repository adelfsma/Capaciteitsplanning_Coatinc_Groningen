
import os
from datetime import date
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
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
    PRODUCTIE_DASHBOARD_KOLOMMEN,
    previous_workday,
    load_published_data,
    load_metadata,
    load_mis_data,
    load_dashboard_manual,
    save_dashboard_manual_entry,
    build_dashboard_data,
    build_productie_dashboard_week,
    build_productie_dashboard_ytd,
    compute_otif,
    get_peildatum_from_metadata,
    find_originele_leverdatum_column,
    format_int,
    format_pct,
    render_environment_banner,
    get_page_title,
    is_test_environment,
    get_locatie_naam,
    get_locatie_logo,
    get_poetsen_otif_actief,
    get_max_capaciteit_config,
    get_kg_traverse_defaults,
    get_otif_uitsluiten_statussen,
    get_otif_uitsluiten_klanten,
    get_beheer_wachtwoord,
    get_norm_manuren_per_ton,
    get_norm_traversen_per_dag,
    get_norm_gem_gewicht_per_traverse,
)

st.set_page_config(layout="wide", page_title=get_page_title(f"Capaciteitsplanning {get_locatie_naam()}"))

def bereken_aantal_balken(
    day_df: pd.DataFrame,
    kg_per_traverse_constructie: float,
    kg_per_traverse_maatwerk: float,
    kg_per_traverse_seriewerk: float,
) -> pd.DataFrame:
    """Bereken het totale aantal balken per dag op basis van het materiaaltype."""
    result = day_df.copy()

    def materiaal_kg(materiaaltype: str) -> pd.Series:
        col = MATERIAALTYPE_DAG_COLS[materiaaltype]
        if col in result.columns:
            return pd.to_numeric(result[col], errors="coerce").fillna(0.0)
        return pd.Series(0.0, index=result.index)

    result["Aantal_balken_berekend"] = (
        materiaal_kg("Constructie") / kg_per_traverse_constructie
        + materiaal_kg("Maatwerk") / kg_per_traverse_maatwerk
        + materiaal_kg("Seriewerk") / kg_per_traverse_seriewerk
        + materiaal_kg("Overig / onbekend") / kg_per_traverse_maatwerk
    )
    return result


def voeg_aantal_balken_lijn_toe(ax, x, plot_df: pd.DataFrame):
    """Voeg de balkenlijn met waarde-labels en een rechter y-as toe.

    De rode lijn zelf loopt door het volle plotgebied (visueel: trend van
    de dag-belasting). De waarde-labels staan echter in een vaste rij
    onderaan het plotgebied, net boven de x-as. Op die manier kunnen ze
    nooit overlappen met de kg-totalen bovenaan de staven, met de segment-
    labels binnenin de staven, of met de legenda bovenaan. Elk label
    krijgt bovendien een subtiel wit-met-rode-rand kadertje voor extra
    leesbaarheid.
    """
    aantal_balken = plot_df["Aantal_balken_berekend"].fillna(0).astype(float).to_numpy()
    ax_right = ax.twinx()
    lijn, = ax_right.plot(
        x,
        aantal_balken,
        color="#C00000",
        marker="o",
        markersize=4,
        linewidth=2,
        label="Aantal balken",
        zorder=5,
    )
    ax_right.set_ylabel("Aantal balken", color="#C00000")
    ax_right.tick_params(axis="y", colors="#C00000")
    ax_right.spines["top"].set_visible(False)
    ax_right.spines["right"].set_alpha(0.45)
    ax_right.grid(False)

    # Secundaire y-as ruim schalen: de hoogste rode waarde komt op ~40 %
    # van de plot-hoogte. Zo blijven de lijn en zijn markers altijd in de
    # onderste helft van het plotgebied en kunnen ze nooit door de kg-
    # totalen bovenaan de staven heen lopen.
    max_val = float(max(aantal_balken)) if len(aantal_balken) else 0.0
    ymax_right = (max_val / 0.40) if max_val > 0 else 1.0
    ax_right.set_ylim(0, ymax_right)

    # Waarde-labels in een vaste rij onderaan het plotgebied. De
    # x-coördinaat is de dag (data), de y-coördinaat een vaste fractie
    # van de as (get_xaxis_transform combineert dit correct). Doordat
    # alle labels dezelfde y-positie hebben kunnen ze niet meer met
    # andere labels of balken overlappen.
    trans = ax_right.get_xaxis_transform()
    for i, value in enumerate(aantal_balken):
        if value > 0:
            label = f"{value:.1f}".replace(".", ",")
            ax_right.text(
                i,
                0.02,
                label,
                transform=trans,
                ha="center",
                va="bottom",
                fontsize=7,
                fontweight="bold",
                color="#C00000",
                bbox={
                    "boxstyle": "round,pad=0.2",
                    "facecolor": "white",
                    "edgecolor": "#C00000",
                    "linewidth": 0.6,
                    "alpha": 0.95,
                },
                zorder=6,
                clip_on=False,
            )
    return lijn


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

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right", rotation_mode="anchor")
    ax.set_ylabel("KG")
    ax.yaxis.set_major_formatter(
    FuncFormatter(lambda value, position: format_int(value))
    )
    ax.grid(axis="y", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_alpha(0.3)
    ax.spines["bottom"].set_alpha(0.3)
    balken_lijn = voeg_aantal_balken_lijn_toe(ax, x, plot_df)
    handles, legend_labels = ax.get_legend_handles_labels()
    # Legenda boven de plot, tussen titel en grafiek. Voorkomt overlap
    # met balken; ncols laat 'm horizontaal uitspreiden.
    ax.legend(
        handles + [balken_lijn],
        legend_labels + ["Aantal balken"],
        loc="lower center",
        bbox_to_anchor=(0.5, 1.02),
        ncols=len(handles) + 1,
        frameon=False,
        fontsize=9,
    )

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
    # Extra top-marge zodat de legenda tussen titel en plot past.
    fig.tight_layout(rect=(0, 0, 1, 0.92))
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
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right", rotation_mode="anchor")
    ax.set_ylabel("KG")
    ax.yaxis.set_major_formatter(
    FuncFormatter(lambda value, position: format_int(value))
    )
    ax.grid(axis="y", alpha=0.25)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_alpha(0.3)
    ax.spines["bottom"].set_alpha(0.3)
    balken_lijn = voeg_aantal_balken_lijn_toe(ax, x, plot_df)
    handles, legend_labels = ax.get_legend_handles_labels()
    # Legenda boven de plot, tussen titel en grafiek. Voorkomt overlap met
    # gestapelde balken; ncols=aantal items zodat 't in één rij past.
    ax.legend(
        handles + [balken_lijn],
        legend_labels + ["Aantal balken"],
        loc="lower center",
        bbox_to_anchor=(0.5, 1.02),
        ncols=len(handles) + 1,
        frameon=False,
        fontsize=9,
    )

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

    # Extra top-marge zodat de legenda tussen titel en plot past.
    fig.tight_layout(rect=(0, 0, 1, 0.92))
    return fig


def _otif_pct_to_angle(
    pct: float,
    threshold_green: float = OTIF_GAUGE_GREEN_THRESHOLD,
    threshold_orange: float = OTIF_GAUGE_ORANGE_THRESHOLD,
) -> float:
    """
    Zet een OTIF-percentage om naar een arc-hoek volgens een piecewise-lineaire
    schaal waarbij elke zone (rood, oranje, groen) precies 60° van de
    halfronde meter beslaat. Zo is het groene gebied altijd goed zichtbaar,
    ook al is 96–100 % maar een smalle band in de businessdefinitie.
    De naaldpositie binnen een zone blijft lineair, dus goed afleesbaar.
    """
    clip = max(0.0, min(100.0, pct))
    if clip <= threshold_orange:
        return 180.0 - (clip / threshold_orange) * 60.0
    if clip <= threshold_green:
        return 120.0 - ((clip - threshold_orange) / (threshold_green - threshold_orange)) * 60.0
    return 60.0 - ((clip - threshold_green) / (100.0 - threshold_green)) * 60.0


def _otif_zone_color(
    pct: float | None,
    threshold_green: float = OTIF_GAUGE_GREEN_THRESHOLD,
    threshold_orange: float = OTIF_GAUGE_ORANGE_THRESHOLD,
) -> str:
    """Bepaal de kleur (groen/oranje/rood/grijs) voor een OTIF-waarde."""
    if pct is None:
        return "#6b7280"
    if pct >= threshold_green:
        return "#16a34a"
    if pct >= threshold_orange:
        return "#f59e0b"
    return "#dc2626"


def make_otif_gauge(
    otif_pct: float | None,
    threshold_green: float = OTIF_GAUGE_GREEN_THRESHOLD,
    threshold_orange: float = OTIF_GAUGE_ORANGE_THRESHOLD,
    title: str = "OTIF",
    figsize: tuple[float, float] = (5.2, 3.4),
    show_value: bool = True,
    show_title: bool = True,
):
    """
    Matplotlib-variant van de OTIF-snelheidsmeter voor het OTIF-tabblad.
    - < 80 %             → rood
    - 80 % – < 96 %      → oranje
    - ≥ 96 %             → groen
    De arc is verdeeld in drie even brede zones (elk 60°); binnen elke zone
    is de naaldpositie lineair.
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

    outer_r = 1.0
    ring_width = 0.28

    # Achtergrondring om ontbrekende data netjes te tonen.
    ax.add_patch(Wedge((0, 0), outer_r, 0, 180, width=ring_width, facecolor=GRIJS, edgecolor="none"))

    # Drie even brede zones op vaste posities.
    ax.add_patch(Wedge((0, 0), outer_r, 120, 180, width=ring_width, facecolor=ROOD,   edgecolor="none"))
    ax.add_patch(Wedge((0, 0), outer_r,  60, 120, width=ring_width, facecolor=ORANJE, edgecolor="none"))
    ax.add_patch(Wedge((0, 0), outer_r,   0,  60, width=ring_width, facecolor=GROEN,  edgecolor="none"))

    # Labels op de zonegrenzen + intermediair per zone voor context.
    for pct_val in (0, threshold_orange / 2, threshold_orange,
                    (threshold_orange + threshold_green) / 2, threshold_green,
                    (threshold_green + 100) / 2, 100):
        angle_deg = _otif_pct_to_angle(pct_val, threshold_green, threshold_orange)
        angle_rad = np.deg2rad(angle_deg)
        x1 = (outer_r - ring_width) * np.cos(angle_rad)
        y1 = (outer_r - ring_width) * np.sin(angle_rad)
        x2 = (outer_r - ring_width - 0.06) * np.cos(angle_rad)
        y2 = (outer_r - ring_width - 0.06) * np.sin(angle_rad)
        ax.plot([x1, x2], [y1, y2], color=DONKER, linewidth=1.1)
        xl = (outer_r - ring_width - 0.16) * np.cos(angle_rad)
        yl = (outer_r - ring_width - 0.16) * np.sin(angle_rad)
        # Belangrijkste labels dikker, halverwege-labels wat lichter.
        is_boundary = pct_val in (0, threshold_orange, threshold_green, 100)
        ax.text(
            xl, yl,
            f"{pct_val:.0f}" if pct_val == int(pct_val) else f"{pct_val:.1f}",
            ha="center", va="center",
            fontsize=8 if is_boundary else 7,
            color=DONKER if is_boundary else MIDGRIJS,
            fontweight="bold" if is_boundary else "normal",
        )

    if show_title:
        ax.text(0, 1.14, title, ha="center", va="center", fontsize=13, fontweight="bold", color=DONKER)

    if otif_pct is None:
        ax.plot([0], [0], marker="o", color=MIDGRIJS, markersize=9)
        if show_value:
            ax.text(0, -0.28, "geen data", ha="center", va="center",
                    fontsize=14, fontweight="bold", color=MIDGRIJS)
        return fig

    needle_angle = np.deg2rad(_otif_pct_to_angle(float(otif_pct), threshold_green, threshold_orange))
    needle_len = outer_r - ring_width - 0.02
    nx = needle_len * np.cos(needle_angle)
    ny = needle_len * np.sin(needle_angle)
    ax.plot([0, nx], [0, ny], color=DONKER, linewidth=2.8, solid_capstyle="round")
    ax.plot([0], [0], marker="o", color=DONKER, markersize=11)

    if show_value:
        num_color = _otif_zone_color(otif_pct, threshold_green, threshold_orange)
        ax.text(0, -0.30, f"{otif_pct:.1f}%", ha="center", va="center",
                fontsize=24, fontweight="bold", color=num_color)

    fig.tight_layout()
    return fig


def make_otif_zone_bar_svg(
    otif_pct: float | None,
    threshold_green: float = OTIF_GAUGE_GREEN_THRESHOLD,
    threshold_orange: float = OTIF_GAUGE_ORANGE_THRESHOLD,
    width: int = 260,
    height: int = 18,
) -> str:
    """
    Compacte horizontale 3-zone bar (mini-snelheidsmeter) voor gebruik in de
    OTIF-dashboardkaart. De bar bestaat uit drie gelijke segmenten
    (rood/oranje/groen) en een kleine driehoek als indicator van de huidige
    OTIF-waarde. Blijft ruim binnen de hoogte van een tekstregel, zodat de
    OTIF-kaart precies dezelfde totale hoogte krijgt als de voorraadkaarten.
    """
    ROOD, ORANJE, GROEN = "#dc2626", "#f59e0b", "#16a34a"
    DONKER, MIDGRIJS = "#111827", "#6b7280"

    bar_y = 0
    bar_h = 6
    seg_w = width / 3.0

    def pct_to_x(pct: float) -> float:
        clip = max(0.0, min(100.0, pct))
        if clip <= threshold_orange:
            return (clip / threshold_orange) * seg_w
        if clip <= threshold_green:
            return seg_w + ((clip - threshold_orange) / (threshold_green - threshold_orange)) * seg_w
        return 2 * seg_w + ((clip - threshold_green) / (100.0 - threshold_green)) * seg_w

    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {width} {height}" '
        f'width="100%" style="display:block; max-width:100%; height:auto;" '
        f'aria-label="OTIF zone-indicator">'
    ]
    # Drie zones. Buitenste hoeken licht afgerond voor een verfijnde look.
    parts.append(f'<rect x="0" y="{bar_y}" width="{seg_w:.2f}" height="{bar_h}" '
                 f'fill="{ROOD}" rx="2" ry="2" />')
    parts.append(f'<rect x="{seg_w:.2f}" y="{bar_y}" width="{seg_w:.2f}" height="{bar_h}" '
                 f'fill="{ORANJE}" />')
    parts.append(f'<rect x="{2*seg_w:.2f}" y="{bar_y}" width="{seg_w:.2f}" height="{bar_h}" '
                 f'fill="{GROEN}" rx="2" ry="2" />')

    # Indicator: driehoek onder de bar wijzend naar de huidige waarde.
    if otif_pct is None:
        cx = width / 2.0
        parts.append(
            f'<polygon points="{cx-3.5:.2f},{bar_h+11} {cx+3.5:.2f},{bar_h+11} '
            f'{cx:.2f},{bar_h+3}" fill="{MIDGRIJS}" />'
        )
    else:
        x = pct_to_x(float(otif_pct))
        parts.append(
            f'<polygon points="{x-4:.2f},{bar_h+11} {x+4:.2f},{bar_h+11} '
            f'{x:.2f},{bar_h+3}" fill="{DONKER}" />'
        )

    parts.append("</svg>")
    return "".join(parts)


def make_otif_gauge_svg(
    otif_pct: float | None,
    threshold_green: float = OTIF_GAUGE_GREEN_THRESHOLD,
    threshold_orange: float = OTIF_GAUGE_ORANGE_THRESHOLD,
    width: int = 300,
    height: int = 160,
) -> str:
    """
    Halfronde SVG-snelheidsmeter (niet meer op het dashboard; alleen nog
    beschikbaar mocht er iemand een volledige gauge willen inbouwen). De
    productieweergave gebruikt make_otif_zone_bar_svg voor compactheid.
    """
    import math

    ROOD, ORANJE, GROEN = "#dc2626", "#f59e0b", "#16a34a"
    DONKER, MIDGRIJS = "#111827", "#6b7280"

    cx = width / 2
    cy = height * 0.86
    r_outer = min(width * 0.40, height * 0.72)
    r_inner = r_outer * 0.60
    label_r = r_inner - 12

    def polar(r: float, a_deg: float) -> tuple[float, float]:
        rad = math.radians(a_deg)
        return (cx + r * math.cos(rad), cy - r * math.sin(rad))

    def ring_segment_path(a_start: float, a_end: float, steps: int = 32) -> str:
        outer_pts, inner_pts = [], []
        for i in range(steps + 1):
            t = i / steps
            a = a_start + t * (a_end - a_start)
            outer_pts.append(polar(r_outer, a))
            inner_pts.append(polar(r_inner, a))
        parts = [f"M {outer_pts[0][0]:.2f},{outer_pts[0][1]:.2f}"]
        for x, y in outer_pts[1:]:
            parts.append(f"L {x:.2f},{y:.2f}")
        for x, y in reversed(inner_pts):
            parts.append(f"L {x:.2f},{y:.2f}")
        parts.append("Z")
        return " ".join(parts)

    svg_parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {width} {height}" '
        f'width="100%" style="display:block; max-width:100%; height:auto;" '
        f'aria-label="OTIF snelheidsmeter">'
    ]

    # Zones (elk 60° op de arc).
    svg_parts.append(f'<path d="{ring_segment_path(180, 120)}" fill="{ROOD}" />')
    svg_parts.append(f'<path d="{ring_segment_path(120,  60)}" fill="{ORANJE}" />')
    svg_parts.append(f'<path d="{ring_segment_path( 60,   0)}" fill="{GROEN}" />')

    # Streepjes en labels op de zonegrenzen.
    for pct_val in (0, threshold_orange, threshold_green, 100):
        a = _otif_pct_to_angle(pct_val, threshold_green, threshold_orange)
        x1, y1 = polar(r_inner, a)
        x2, y2 = polar(r_inner - 5, a)
        svg_parts.append(
            f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" '
            f'stroke="{DONKER}" stroke-width="1.2" />'
        )
        xl, yl = polar(label_r, a)
        svg_parts.append(
            f'<text x="{xl:.2f}" y="{yl:.2f}" text-anchor="middle" '
            f'dominant-baseline="central" font-size="10" font-weight="600" '
            f'fill="{DONKER}" font-family="sans-serif">{int(pct_val)}</text>'
        )

    # Naald.
    if otif_pct is None:
        svg_parts.append(
            f'<circle cx="{cx:.2f}" cy="{cy:.2f}" r="6" fill="{MIDGRIJS}" />'
        )
    else:
        a = _otif_pct_to_angle(float(otif_pct), threshold_green, threshold_orange)
        needle_len = r_outer - 4
        nx, ny = polar(needle_len, a)
        svg_parts.append(
            f'<line x1="{cx:.2f}" y1="{cy:.2f}" x2="{nx:.2f}" y2="{ny:.2f}" '
            f'stroke="{DONKER}" stroke-width="3" stroke-linecap="round" />'
        )
        svg_parts.append(
            f'<circle cx="{cx:.2f}" cy="{cy:.2f}" r="7" fill="{DONKER}" />'
        )

    svg_parts.append("</svg>")
    return "".join(svg_parts)


_logo = get_locatie_logo()
if os.path.exists(_logo):
    st.sidebar.image(_logo, width=200)
st.sidebar.caption(APP_VERSION)

# Handmatige cache-refresh. Data uit de cloud wordt automatisch elke 60s
# ververst; deze knop forceert een directe refresh na een nieuwe publicatie.
if st.sidebar.button("🔄 Ververs data", help="Haal de nieuwste data uit de cloud"):
    st.cache_data.clear()
    st.rerun()

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

if is_test_environment():
    st.markdown(
        f"<h1 style='margin-bottom:0'>Capaciteitsplanning {get_locatie_naam()}"
        f"&nbsp;<span style='"
        f"background:#dc2626;color:white;font-size:0.55em;font-weight:800;"
        f"padding:4px 14px;border-radius:8px;vertical-align:middle;"
        f"letter-spacing:0.08em;text-transform:uppercase'>TEST</span></h1>",
        unsafe_allow_html=True,
    )
else:
    st.title(f"Capaciteitsplanning {get_locatie_naam()}")
st.sidebar.header("Instellingen")
_cap_min, _cap_max, _cap_default, _cap_step = get_max_capaciteit_config()
capaciteit_ton = st.sidebar.slider("Max capaciteit per dag (ton)", _cap_min, _cap_max, _cap_default, _cap_step)
capaciteit_kg = capaciteit_ton * 1000
offset = st.sidebar.selectbox("Verzinkdatum = leverdatum - X werkdagen", [1, 2, 3, 4], index=1)
_kg_constr_def, _kg_maat_def, _kg_serie_def = get_kg_traverse_defaults()
kg_per_traverse_constructie = st.sidebar.number_input("KG per traverse Constructie", min_value=100, max_value=10000, value=_kg_constr_def, step=50)
kg_per_traverse_maatwerk = st.sidebar.number_input("KG per traverse Maatwerk", min_value=100, max_value=10000, value=_kg_maat_def, step=50)
kg_per_traverse_seriewerk = st.sidebar.number_input("KG per traverse Seriewerk", min_value=100, max_value=10000, value=_kg_serie_def, step=50)
default_start = previous_workday(date.today())
startdatum = st.sidebar.date_input("Startdatum rapport", value=default_start)
toon_alle_regels = st.sidebar.checkbox("Toon alle regels in controletab", value=True)

try:
    df_raw, export_file_summary, order_file, holiday_df = load_published_data()
except Exception as e:
    st.error(f"Kan gepubliceerde data niet laden: {e}")
    st.stop()

df, df_plan, dag, week, advies_datum = build_dashboard_data(
    df_raw,
    holiday_df,
    startdatum,
    capaciteit_kg,
    offset,
    kg_per_traverse_constructie,
)
dag = bereken_aantal_balken(
    dag,
    kg_per_traverse_constructie,
    kg_per_traverse_maatwerk,
    kg_per_traverse_seriewerk,
)

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

# Feestdagen doorgeven zodat 'vorige werkdag' (voor de depot-shift) feestdagen
# overslaat. holiday_df bevat de datums uit de metadata-feestdagenkalender.
otif_holiday_dates = set(holiday_df["Datum"].tolist()) if not holiday_df.empty else set()

otif_result = compute_otif(
    df,
    peildatum=otif_peildatum,
    date_column=otif_date_column,
    date_column_label=otif_date_label,
    holiday_dates=otif_holiday_dates,
    poetsen_afgehaald_niet_ok=get_poetsen_otif_actief(),
    uitsluiten_statussen=get_otif_uitsluiten_statussen(),
    uitsluiten_klanten=get_otif_uitsluiten_klanten(),
)

tab1, tab_prod, tab2, tab3, tab4 = st.tabs(
    ["Voorraad & OTIF", "Productie dashboard", "Gebruikte gegevens", "OTIF", "Debug"]
)

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

    # Gedeelde opmaak voor de drie kaarten. De inhoud van elke kaart heeft
    # dezelfde structuur (label + waarde + detailregel), waardoor Streamlit
    # ze automatisch op gelijke hoogte rendert.
    st.markdown(
        """
        <style>
        .cgr-stock-card {
            border: 1px solid #d0d7de;
            border-radius: 12px;
            padding: 0.85rem 1.05rem 0.9rem 1.05rem;
            background: #ffffff;
            box-shadow: 0 1px 2px rgba(16, 24, 40, 0.05);
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            gap: 0.35rem;
            height: 100%;
        }
        .cgr-stock-card-accent-black { border-left: 6px solid #1f2937; }
        .cgr-stock-card-accent-white { border-left: 6px solid #94a3b8; }
        .cgr-stock-card-accent-otif  { border-left: 6px solid #16a34a; }
        .cgr-stock-label {
            font-size: 0.86rem;
            color: #475569;
            font-weight: 600;
            letter-spacing: 0.005em;
        }
        .cgr-stock-value {
            font-size: 1.75rem;
            line-height: 1.05;
            color: #111827;
            font-weight: 700;
        }
        .cgr-stock-zonebar {
            margin: 0.05rem 0 0.1rem 0;
            max-width: 220px;
        }
        .cgr-stock-detail {
            font-size: 0.82rem;
            color: #64748b;
            line-height: 1.45;
            margin-top: auto;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col_zwart, col_wit, col_otif = st.columns(3)

    with col_zwart:
        st.markdown(
            f"""
            <div class="cgr-stock-card cgr-stock-card-accent-black">
                <div class="cgr-stock-label">Zwarte voorraad (kg)</div>
                <div class="cgr-stock-value">{format_int(kg_totale_zwarte_voorraad)}</div>
                <div class="cgr-stock-detail">
                    Productie gereed: <strong>{format_int(kg_productie_gereed)}</strong> kg &nbsp;·&nbsp;
                    binnengemeld / voorbewerking / geblokkeerd: <strong>{format_int(kg_voorbewerking_geblokkeerd)}</strong> kg
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
        # OTIF: identieke kaartstructuur als voorraadkaarten (label + waarde
        # + detailregel), aangevuld met een subtiele 3-zone bar als
        # miniatuur-snelheidsmeter. De volledige gauge staat in het
        # OTIF-tabblad voor drill-down.
        peildatum_str = pd.Timestamp(otif_result["peildatum"]).strftime("%d-%m-%Y")
        otif_num_color = _otif_zone_color(otif_result["otif_pct"])
        otif_pct_html = (
            "geen data" if otif_result["otif_pct"] is None
            else f"{otif_result['otif_pct']:.1f}%"
        )
        if otif_result["totaal"] > 0:
            samenvatting = (
                f"{otif_result['gereed']} op tijd · "
                f"{otif_result['niet_gereed']} te laat · "
                f"{otif_result['nvt']} nvt"
            )
            if otif_result["onbekend"] > 0:
                samenvatting += f" · {otif_result['onbekend']} onbekend"
            samenvatting += f" (totaal {otif_result['totaal']})"
            
        else:
            samenvatting = "Geen orders met deze peildatum in de dataset."

        zonebar_svg = make_otif_zone_bar_svg(otif_result["otif_pct"])

        st.markdown(
            f"""
            <div class="cgr-stock-card cgr-stock-card-accent-otif">
                <div class="cgr-stock-label">OTIF · {otif_result['date_column_label']} = {peildatum_str}</div>
                <div class="cgr-stock-value" style="color: {otif_num_color};">{otif_pct_html}</div>
                <div class="cgr-stock-zonebar">{zonebar_svg}</div>
                <div class="cgr-stock-detail">{samenvatting}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.subheader("Eerstvolgende leverdatum")
    st.markdown(f'<div style="padding: 1rem 1.25rem; border-radius: 12px; border: 1px solid #d0d7de; background-color: #f6f8fa; margin-bottom: 0.75rem;"><div style="font-size: 2.2rem; font-weight: 700;">{advies_datum.strftime("%d-%m-%Y")}</div></div>', unsafe_allow_html=True)

    st.subheader("Capaciteit versus dagbelasting op verzinkdatum")
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

# ── Tab: Productie dashboard ───────────────────────────────────────────────────
with tab_prod:
    st.subheader("Productie dashboard")
    st.caption(
        "KPI-overzicht per week (bron: MIS-export + capaciteitsplanning). "
        "Historische plan-waarden komen beschikbaar in een volgende versie."
    )

    # MIS-data en handmatige invoer laden
    try:
        mis_df = load_mis_data()
    except Exception as e:
        st.error(f"MIS-bestand kon niet worden ingeladen: {e}")
        mis_df = pd.DataFrame()

    manual_dict = load_dashboard_manual()

    if mis_df.empty:
        st.warning(
            "Er is nog geen MIS-bestand (`mis.xlsx`) beschikbaar in de cloud. "
            "Vraag de beheerder om deze te uploaden via de beheeromgeving."
        )

    # Week-selector: huidige ISO-week als default
    _today = pd.Timestamp(date.today()).normalize()
    _iso_now = _today.isocalendar()
    c_year, c_week, _c_spacer = st.columns([1, 1, 4])
    with c_year:
        gekozen_jaar = st.number_input(
            "Jaar", min_value=2024, max_value=2099, value=int(_iso_now.year), step=1
        )
    with c_week:
        gekozen_week = st.number_input(
            "Weeknummer", min_value=1, max_value=53, value=int(_iso_now.week), step=1
        )

    _feest_set = set(pd.to_datetime(holiday_df["Datum"]).tolist()) if not holiday_df.empty else set()

    # Bouw beide weken (huidige/gekozen + vorige)
    _monday_current = pd.Timestamp.fromisocalendar(int(gekozen_jaar), int(gekozen_week), 1)
    _monday_prev = _monday_current - pd.Timedelta(days=7)
    _iso_prev = _monday_prev.isocalendar()

    week_curr = build_productie_dashboard_week(
        mis_df, dag, manual_dict, int(gekozen_jaar), int(gekozen_week), _feest_set,
    )
    week_prev = build_productie_dashboard_week(
        mis_df, dag, manual_dict, int(_iso_prev.year), int(_iso_prev.week), _feest_set,
    )

    # ── Handmatige invoer (wachtwoordbeveiligd) ────────────────────────────────
    with st.expander("🔒 Handmatige velden invoeren (beheer)"):
        pw = st.text_input("Beheerwachtwoord", type="password", key="prod_dashboard_pw")
        if pw == "":
            st.caption("Vul het beheerwachtwoord in om handmatige velden op te slaan.")
        elif pw != get_beheer_wachtwoord():
            st.error("Onjuist wachtwoord.")
        else:
            st.success("Ingelogd als beheerder.")
            invoer_datum = st.date_input(
                "Datum",
                value=date.today(),
                help="Kies de dag waarop de handmatige velden betrekking hebben.",
            )
            bestaand = manual_dict.get(pd.Timestamp(invoer_datum).strftime("%Y-%m-%d"), {})
            c1, c2, c3 = st.columns(3)
            with c1:
                inp_klachten = st.number_input(
                    "Aantal klachten",
                    min_value=0, step=1,
                    value=int(bestaand.get("klachten", 0)),
                )
            with c2:
                inp_storing = st.number_input(
                    "Storingstijd (min)",
                    min_value=0, step=1,
                    value=int(bestaand.get("storingstijd_min", 0)),
                )
            with c3:
                inp_veilig = st.number_input(
                    "Veiligheidsincidenten",
                    min_value=0, step=1,
                    value=int(bestaand.get("veiligheidsincidenten", 0)),
                )
            if st.button("Opslaan", type="primary"):
                try:
                    save_dashboard_manual_entry(invoer_datum, inp_klachten, inp_storing, inp_veilig)
                    st.success(f"✅ Opgeslagen voor {invoer_datum.strftime('%d-%m-%Y')}.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Opslaan mislukt: {e}")

    # ── Weektabellen onder elkaar (compact zodat ze passen zonder scrollen) ────
    # Kortere kolomlabels + kleinere breedtes zodat de tabel binnen de container past.
    _PROD_COL_CONFIG = {
        "Dag":                    st.column_config.TextColumn("Dag",       width="small"),
        "Datum":                  st.column_config.TextColumn("Datum",     width="small"),
        "TONNAGE PLAN":           st.column_config.TextColumn("Plan (kg)", width="small"),
        "TONNAGE WERKELIJK":      st.column_config.TextColumn("Werkelijk (kg)", width="small"),
        "MANUREN / TON":          st.column_config.TextColumn("mu/ton",    width="small"),
        "AFKEUR IN KG":           st.column_config.TextColumn("Afkeur (kg)", width="small"),
        "AANTAL TRAVERSEN":       st.column_config.TextColumn("Trav.",     width="small"),
        "GEM GEWICHT PER TR":     st.column_config.TextColumn("kg/trav.",  width="small"),
        "AANTAL KLACHTEN":        st.column_config.TextColumn("Klachten",  width="small"),
        "Storingstijd in min":    st.column_config.TextColumn("Storing (min)", width="small"),
        "VEILIGHEIDSINCIDENTEN":  st.column_config.TextColumn("Veiligheid", width="small"),
    }

    def _fmt_week_df(week_df: pd.DataFrame) -> pd.DataFrame:
        """Format datums als dd-mm en getallen met NL duizendtal-punt."""
        out = week_df.copy()
        out["Datum"] = pd.to_datetime(out["Datum"]).dt.strftime("%d-%m")
        int_cols = ["TONNAGE PLAN", "TONNAGE WERKELIJK", "AFKEUR IN KG",
                    "AANTAL TRAVERSEN", "GEM GEWICHT PER TR", "AANTAL KLACHTEN",
                    "Storingstijd in min", "VEILIGHEIDSINCIDENTEN"]
        for c in int_cols:
            if c in out.columns:
                out[c] = out[c].apply(lambda v: format_int(v) if pd.notna(v) else "")
        if "MANUREN / TON" in out.columns:
            out["MANUREN / TON"] = out["MANUREN / TON"].apply(
                lambda v: f"{v:.2f}".replace(".", ",") if pd.notna(v) else ""
            )
        return out

    st.markdown(f"##### Week {int(gekozen_week)} — huidige selectie")
    st.dataframe(
        _fmt_week_df(week_curr),
        width="stretch", hide_index=True,
        column_config=_PROD_COL_CONFIG,
    )

    st.markdown(f"##### Week {int(_iso_prev.week)} — vorige week")
    st.dataframe(
        _fmt_week_df(week_prev),
        width="stretch", hide_index=True,
        column_config=_PROD_COL_CONFIG,
    )

    # ── Chart helpers (verfijnde stijl) ────────────────────────────────────────
    # Kleuren afgestemd op de bestaande grafieken in de app (blauw voor werkelijk,
    # subtiel rood voor de norm-lijn).
    _KLEUR_WERKELIJK = "#2E75B6"
    _KLEUR_NORM      = "#C00000"
    _KLEUR_GRID      = "#E5E5E5"
    _KLEUR_TEXT      = "#333333"

    def _style_axis(ax):
        """Uniforme, opgeruimde stijl voor alle KPI-grafieken."""
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(_KLEUR_GRID)
        ax.spines["bottom"].set_color(_KLEUR_GRID)
        ax.tick_params(colors=_KLEUR_TEXT, labelsize=9)
        ax.yaxis.grid(True, color=_KLEUR_GRID, linestyle="-", linewidth=0.6, alpha=0.7)
        ax.set_axisbelow(True)
        ax.title.set_color(_KLEUR_TEXT)
        ax.title.set_fontsize(11)
        ax.title.set_fontweight("semibold")
        ax.title.set_loc = "left"
        for spine in ax.spines.values():
            spine.set_linewidth(0.8)

    def _pad_ylim(ax, values, norm_values=None, extra=0.35, floor=0):
        """Zet y-limits op basis van max-waarde met extra kopruimte.
        35% padding voorkomt dat de legenda over de hoge bars valt."""
        candidates = [v for v in values if v is not None and pd.notna(v) and v > 0]
        if norm_values is not None:
            candidates += [v for v in norm_values if v is not None and pd.notna(v) and v > 0]
        if not candidates:
            return
        top = max(candidates) * (1 + extra)
        ax.set_ylim(bottom=floor, top=top)

    def _draw_week_chart(ax, plot_df, kpi_col, norm_getter, y_label, titel):
        """Werkelijk als bars, norm als horizontale streepjeslijn, waarde-labels bovenop."""
        werkelijk = plot_df[kpi_col].tolist()
        if not any(pd.notna(v) and v > 0 for v in werkelijk):
            ax.text(0.5, 0.5, "Geen data in deze week", ha="center", va="center",
                    transform=ax.transAxes, color="#999", fontsize=10)
            ax.set_axis_off()
            ax.set_title(titel, loc="left", pad=10, color=_KLEUR_TEXT,
                         fontsize=11, fontweight="semibold")
            return
        norm_waarden = [norm_getter(d) for d in pd.to_datetime(plot_df["Datum"])]
        x = np.arange(len(plot_df))
        bars = [v if pd.notna(v) else 0 for v in werkelijk]

        ax.bar(x, bars, width=0.60, color=_KLEUR_WERKELIJK, alpha=0.90,
               edgecolor="none", label="Werkelijk", zorder=2)
        ax.plot(x, norm_waarden, color=_KLEUR_NORM, linestyle="--", linewidth=1.6,
                marker="", label="Norm", zorder=3)

        # Waarde-labels bovenaan iedere bar
        for xi, v in zip(x, werkelijk):
            if pd.notna(v) and v > 0:
                label = f"{v:.1f}".replace(".", ",") if v < 100 else format_int(v)
                ax.text(xi, v, label, ha="center", va="bottom", fontsize=8,
                        color=_KLEUR_TEXT, zorder=4)

        ax.set_xticks(x)
        ax.set_xticklabels(
            [f"{r['Dag']}\n{pd.to_datetime(r['Datum']).strftime('%d-%m')}"
             for _, r in plot_df.iterrows()],
            fontsize=9, color=_KLEUR_TEXT,
        )
        ax.set_ylabel(y_label, fontsize=9, color=_KLEUR_TEXT)
        ax.legend(loc="upper right", fontsize=8, frameon=False)
        _style_axis(ax)
        _pad_ylim(ax, werkelijk, norm_waarden)
        ax.set_title(titel, loc="left", pad=10, color=_KLEUR_TEXT,
                     fontsize=11, fontweight="semibold")

    def _draw_ytd_chart(ax, ytd_df, value_col, norm_getter, y_label, titel):
        """YTD per ISO-week: werkelijk als bars, norm als lijn."""
        if ytd_df.empty:
            ax.text(0.5, 0.5, "Geen YTD-data beschikbaar", ha="center", va="center",
                    transform=ax.transAxes, color="#999", fontsize=10)
            ax.set_axis_off()
            ax.set_title(titel, loc="left", pad=10, color=_KLEUR_TEXT,
                         fontsize=11, fontweight="semibold")
            return
        weken = ytd_df["Weeknr"].tolist()
        werkelijk = ytd_df[value_col].tolist()
        norm_waarden = [norm_getter(d) for d in ytd_df["Week_startdatum"]]
        x = np.arange(len(weken))

        ax.bar(x, [v if pd.notna(v) else 0 for v in werkelijk],
               width=0.75, color=_KLEUR_WERKELIJK, alpha=0.85,
               edgecolor="none", label="Werkelijk", zorder=2)
        ax.plot(x, norm_waarden, color=_KLEUR_NORM, linestyle="--", linewidth=1.6,
                marker="", label="Norm", zorder=3)

        # Tick spacing: max ~10 labels om overlap te voorkomen
        step = 1 if len(weken) <= 10 else max(1, len(weken) // 10)
        tick_idx = list(range(0, len(weken), step))
        # Zorg dat de laatste week zichtbaar is — vervang de laatste als hij te
        # dicht op de nieuwe laatste zou komen te staan, anders toevoegen.
        last_i = len(weken) - 1
        if tick_idx and tick_idx[-1] != last_i:
            if last_i - tick_idx[-1] < step:
                tick_idx[-1] = last_i
            else:
                tick_idx.append(last_i)
        ax.set_xticks(tick_idx)
        ax.set_xticklabels([f"w{weken[i]}" for i in tick_idx],
                           fontsize=9, color=_KLEUR_TEXT)
        ax.set_ylabel(y_label, fontsize=9, color=_KLEUR_TEXT)
        ax.legend(loc="upper right", fontsize=8, frameon=False)
        _style_axis(ax)
        _pad_ylim(ax, werkelijk, norm_waarden)
        ax.set_title(titel, loc="left", pad=10, color=_KLEUR_TEXT,
                     fontsize=11, fontweight="semibold")

    # ── Week-grafieken (huidige selectie, ma t/m vr) ───────────────────────────
    st.markdown(f"### Verloop deze week (week {int(gekozen_week)}, ma t/m vr)")
    _plot_week = week_curr.head(5).copy()

    fig_w, axes_w = plt.subplots(1, 3, figsize=(15, 3.6), dpi=110)
    fig_w.patch.set_facecolor("white")
    _draw_week_chart(axes_w[0], _plot_week, "MANUREN / TON",
                     get_norm_manuren_per_ton, "manuren / ton",
                     "Manuren per ton")
    _draw_week_chart(axes_w[1], _plot_week, "AANTAL TRAVERSEN",
                     get_norm_traversen_per_dag, "traversen",
                     "Aantal traversen")
    _draw_week_chart(axes_w[2], _plot_week, "GEM GEWICHT PER TR",
                     get_norm_gem_gewicht_per_traverse, "kg / traverse",
                     "Gemiddeld gewicht per traverse")
    fig_w.tight_layout()
    st.pyplot(fig_w)
    plt.close(fig_w)

    # ── Week-grafieken (vorige week, ma t/m vr) ────────────────────────────────
    st.markdown(f"### Verloop vorige week (week {int(_iso_prev.week)}, ma t/m vr)")
    _plot_week_prev = week_prev.head(5).copy()

    fig_wp, axes_wp = plt.subplots(1, 3, figsize=(15, 3.6), dpi=110)
    fig_wp.patch.set_facecolor("white")
    _draw_week_chart(axes_wp[0], _plot_week_prev, "MANUREN / TON",
                     get_norm_manuren_per_ton, "manuren / ton",
                     "Manuren per ton")
    _draw_week_chart(axes_wp[1], _plot_week_prev, "AANTAL TRAVERSEN",
                     get_norm_traversen_per_dag, "traversen",
                     "Aantal traversen")
    _draw_week_chart(axes_wp[2], _plot_week_prev, "GEM GEWICHT PER TR",
                     get_norm_gem_gewicht_per_traverse, "kg / traverse",
                     "Gemiddeld gewicht per traverse")
    fig_wp.tight_layout()
    st.pyplot(fig_wp)
    plt.close(fig_wp)

    # ── YTD-grafieken (per ISO-week van gekozen jaar) ──────────────────────────
    st.markdown(f"### Year to date ({int(gekozen_jaar)}, per week)")
    ytd_df = build_productie_dashboard_ytd(mis_df, int(gekozen_jaar))
    if ytd_df.empty:
        st.info("Nog geen YTD-data voor dit jaar beschikbaar.")
    else:
        fig_y, axes_y = plt.subplots(1, 3, figsize=(15, 3.6), dpi=110)
        fig_y.patch.set_facecolor("white")
        _draw_ytd_chart(axes_y[0], ytd_df, "Manuren_per_ton_gewogen",
                        get_norm_manuren_per_ton, "manuren / ton",
                        "Manuren per ton (gewogen)")
        _draw_ytd_chart(axes_y[1], ytd_df, "Traversen_totaal",
                        lambda d: get_norm_traversen_per_dag(d) * 5,
                        "traversen / week",
                        "Aantal traversen per week")
        _draw_ytd_chart(axes_y[2], ytd_df, "Gem_gewicht_per_traverse",
                        get_norm_gem_gewicht_per_traverse, "kg / traverse",
                        "Gemiddeld gewicht per traverse")
        fig_y.tight_layout()
        st.pyplot(fig_y)
        plt.close(fig_y)
        st.caption(
            "YTD-aggregatie: manuren/ton en gem gewicht/traverse zijn gewogen naar "
            "kg-productie per dag. Norm voor traversen is dag-norm × 5 werkdagen."
        )


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
    _caption_regels = [
        f"Peildatum: **{peildatum_str}** · vergeleken met **{otif_result['date_column_label']}** · "
        f"reserveringen worden niet meegeteld."
    ]
    if otif_result.get("depot_shift_actief"):
        vw = otif_result.get("peildatum_vorige_werkdag")
        vw_str = pd.Timestamp(vw).strftime("%d-%m-%Y") if vw is not None else ""
        _caption_regels.append(
            f"Depot-orders (`Aanleveren depot = 1`) hebben een deadline van "
            f"17:00 op hun leverdatum en worden pas de eerstvolgende werkdag "
            f"gemeten. Op deze peildatum tellen depot-orders met leverdatum "
            f"**{vw_str}** mee ({otif_result.get('aantal_depot', 0)} stuks)."
        )
    n_ub = otif_result.get("aantal_ub_uitgesloten", 0)
    if n_ub > 0:
        _caption_regels.append(
            f"UB-orders (uitbesteed) zijn uit deze OTIF weggelaten: "
            f"**{n_ub}** order{'s' if n_ub != 1 else ''}, ongeacht of ze als "
            f"'Ja' of 'Nee' zouden zijn geteld. UB hoort niet bij de "
            f"verzinkstraat."
        )
    n_klant = otif_result.get("aantal_klant_uitgesloten", 0)
    if n_klant > 0:
        _caption_regels.append(
            f"Orders van uitgesloten klanten zijn uit deze OTIF weggelaten: "
            f"**{n_klant}** order{'s' if n_klant != 1 else ''}."
        )
    n_coat = otif_result.get("aantal_coat_uitgesloten", 0)
    if n_coat > 0:
        _caption_regels.append(
            f"Coat-orders zijn uit deze OTIF weggelaten: **{n_coat}** order"
            f"{'s' if n_coat != 1 else ''} viel(en) op deze peildatum maar "
            f"hebben een ordernummer van de vorm `…C…` (coat-alleen) of "
            f"eindigen op `-C` (coat-deel van een gecombineerde verzink+coat-"
            f"order). Deze horen niet bij de verzinkstraat en zouden de "
            f"OTIF-score onterecht beïnvloeden."
        )
    st.caption(" ".join(_caption_regels))

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

*De schaal is bewust niet-lineair: elke zone (rood, oranje, groen) beslaat een even groot deel van de meter, zodat de groene zone (96–100 %) duidelijk zichtbaar blijft. Binnen elke zone is de naaldpositie wél lineair.*
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
        # Is_depot (van compute_otif) wordt naast Aanleveren depot getoond
        # zodat direct duidelijk is welke regels een verschoven peildatum
        # hebben.
        otif_cols = [
            "Bronbestand", "Bron_week", "Nummer", "Ordernummer_base",
            "Debiteurnummer", "Klantnaam",
            "Segment_debtor_export", "Materiaaltype",
            "Datum", "Leverdatum", "Originele_leverdatum",
            "Aanleveren depot", "Is_depot",
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
        {"Categorie":"Instellingen","Omschrijving":"KG per traverse Constructie","Waarde":int(kg_per_traverse_constructie)},
        {"Categorie":"Instellingen","Omschrijving":"KG per traverse Maatwerk","Waarde":int(kg_per_traverse_maatwerk)},
        {"Categorie":"Instellingen","Omschrijving":"KG per traverse Seriewerk","Waarde":int(kg_per_traverse_seriewerk)},
        {"Categorie":"Bestanden","Omschrijving":"Orderbestand","Waarde":order_file},
        {"Categorie":"Kalender","Omschrijving":"Aantal feestdagen / sluitingen","Waarde":int(len(holiday_df))},
        {"Categorie":"Records","Omschrijving":"Niet verzinkt","Waarde":int((df["Verzinkstatus"] == "Niet verzinkt").sum())},
        {"Categorie":"Records","Omschrijving":"Voor startdatum rapport","Waarde":int(((df["Verzinkstatus"] == "Niet verzinkt") & (df["Verzinkdatum"] < pd.Timestamp(startdatum).normalize())).sum())},
        {"Categorie":"Records","Omschrijving":"Coat-orders uitgesloten","Waarde":int((df["Reden_uitsluiting"] == "Coat-order").sum())},
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
