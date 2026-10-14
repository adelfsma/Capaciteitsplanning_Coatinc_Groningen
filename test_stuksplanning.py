"""
Tests voor de stuksplanning (v2.5.23): aantal te verzinken stuks per
verzinkdag naast de kg-planning. Standaard uit; alleen actief met
[locatie] stuks_planning = true. De tests raken alleen pure helpers.
"""
import sys
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))

from shared import (  # noqa: E402
    ORDERREGEL_FILE,
    STUKS_BRON_DEFAULT,
    STUKS_BRON_KLANT,
    STUKS_BRON_ORDERREGEL,
    _add_stuks_file,
    bereken_kg_per_stuk_per_klant,
    build_dashboard_data,
    get_kg_per_stuk_default,
    get_stuks_capaciteit_config,
    get_stuks_planning_actief,
    koppel_stuks,
    stuks_uit_orderregels,
)


def _orderregels() -> pd.DataFrame:
    return pd.DataFrame({
        "OrderID": [100, 100, 100, 200, 300, 300],
        "Subbonnummer V": ["100V1", "100V1", "100V2", "200VC1-V", "300V1", None],
        "aantal": [10, 5, 7, 40, 20, 99],
    })


def test_defaults_stuksplanning_uit():
    assert get_stuks_planning_actief() is False
    mn, mx, default, step = get_stuks_capaciteit_config()
    assert (mn, mx, default, step) == (1000, 8000, 3200, 100)
    assert get_kg_per_stuk_default() == 34


def test_orderregels_bestand_alleen_bij_actief():
    req, opt = _add_stuks_file(["OrderExport2G.xlsx"], ["mis.xlsx"], False)
    assert ORDERREGEL_FILE not in opt
    req, opt = _add_stuks_file(["OrderExport2G.xlsx"], ["mis.xlsx"], True)
    assert opt[-1] == ORDERREGEL_FILE and ORDERREGEL_FILE not in req


def test_stuks_per_deelorder_en_order():
    per_deel, per_order = stuks_uit_orderregels(_orderregels())
    assert per_deel["100V1"] == 15
    assert per_deel["100V2"] == 7
    assert per_deel["200VC1-V"] == 40
    # Regels zonder Subbonnummer V (coat-deel) tellen niet mee
    assert per_deel["300V1"] == 20
    assert per_order[300] == 20


def test_klantratio_alleen_gewogen_en_min_deelorders():
    df = pd.DataFrame({
        "Klantnaam": ["A", "A", "A", "B", "B", "C", "C", "C"],
        "Gewicht_export_kg": [100, 300, 200, 50, 60, 0, 0, 0],   # C niet gewogen
    })
    stuks = pd.Series([10, 10, 10, 5, 5, 10, 10, 10])
    ratio = bereken_kg_per_stuk_per_klant(df, stuks)
    assert ratio["a"] == 20.0          # mediaan van 10, 30, 20
    assert "b" not in ratio.index      # maar 2 deelorders
    assert "c" not in ratio.index      # geen gewogen gewicht


def _merged() -> pd.DataFrame:
    rows = []
    # Klant A: drie gewogen deelorders met orderregels → klantratio 20 kg/stuk
    for i, kg in enumerate([200, 200, 200]):
        rows.append({"Nummer": f"10{i}V1", "Klantnaam": "A", "Gewicht_export_kg": kg,
                     "Gewicht_effectief_kg": kg, "Gewicht_bron": "Export+"})
    # Reservering klant A (geen orderregels) → 1000 / 20 = 50 stuks
    rows.append({"Nummer": "900", "Klantnaam": "A", "Gewicht_export_kg": np.nan,
                 "Gewicht_effectief_kg": 1000, "Gewicht_bron": "Reservering"})
    # Reservering onbekende klant → 680 / 34 = 20 stuks
    rows.append({"Nummer": "901", "Klantnaam": "Z", "Gewicht_export_kg": np.nan,
                 "Gewicht_effectief_kg": 680, "Gewicht_bron": "Reservering"})
    # Reservering zonder gewicht → 0 stuks
    rows.append({"Nummer": "902", "Klantnaam": "Z", "Gewicht_export_kg": np.nan,
                 "Gewicht_effectief_kg": 0, "Gewicht_bron": "Reservering"})
    return pd.DataFrame(rows)


def _regels_voor_merged() -> pd.DataFrame:
    return pd.DataFrame({
        "OrderID": [100, 101, 102, 900],
        "Subbonnummer V": ["100V1", "101V1", "102V1", "900"],   # 900 = reservering-nummer
        "aantal": [10, 10, 10, 5000],
    })


def test_koppel_stuks_werkelijk_en_geschat():
    out = koppel_stuks(_merged(), _regels_voor_merged(), 34)
    assert list(out["Stuks"]) == [10, 10, 10, 50, 20, 0]
    assert list(out["Stuks_bron"]) == [STUKS_BRON_ORDERREGEL] * 3 + [
        STUKS_BRON_KLANT, STUKS_BRON_DEFAULT, STUKS_BRON_DEFAULT]
    # Reserveringen worden nooit aan orderregels gekoppeld, ook niet bij gelijk nummer
    assert out.loc[3, "Stuks"] == 50
    assert np.isnan(out.loc[0, "Kg_per_stuk_gebruikt"])
    assert out.loc[3, "Kg_per_stuk_gebruikt"] == 20.0


def test_koppel_stuks_zonder_orderregels():
    out = koppel_stuks(_merged(), None, 34)
    assert (out["Stuks_bron"] != STUKS_BRON_ORDERREGEL).all()
    assert out.loc[0, "Stuks"] == round(200 / 34)


def _dashboard_input() -> pd.DataFrame:
    d = pd.Timestamp("2026-10-19")   # maandag
    return pd.DataFrame({
        "Nummer": ["1V1", "2V1", "3"],
        "Status": ["Gelost", "Gelost", "Reservering"],
        "Verzinkstatus": ["Niet verzinkt"] * 3,
        "Leverdatum": [d, d, d],
        "Gewicht_effectief_kg": [1000.0, 2000.0, 500.0],
        "Gewicht_bron": ["Export+", "Export+", "Reservering"],
        "Stuks": [100.0, 300.0, 50.0],
        "Stuks_bron": [STUKS_BRON_ORDERREGEL, STUKS_BRON_ORDERREGEL, STUKS_BRON_DEFAULT],
    })


def test_dashboard_dag_en_week_met_stuks():
    holiday = pd.DataFrame(columns=["Datum", "Omschrijving", "Type"])
    _, _, dag, week, _ = build_dashboard_data(
        _dashboard_input(), holiday, date(2026, 10, 12), 60000, 2, 2000,
        capaciteit_stuks=300,
    )
    rij = dag[dag["Verzinkdatum"] == pd.Timestamp("2026-10-15")].iloc[0]
    assert rij["Stuks"] == 450
    assert rij["Stuks_definitief"] == 400
    assert rij["Stuks_reservering"] == 50
    assert rij["Stuks_geschat"] == 50
    assert rij["Capaciteit_stuks"] == 300
    assert rij["Benutting_stuks_pct"] == 150
    assert rij["Status_stuks"] == "🔴"
    assert rij["Kg_per_stuk"] == 3500 / 450
    leeg = dag[dag["Verzinkdatum"] == pd.Timestamp("2026-10-14")].iloc[0]
    assert leeg["Stuks"] == 0 and leeg["Status_stuks"] == "🟢"
    assert {"Stuks", "Capaciteit_stuks", "Benutting_stuks_pct", "Status_stuks"} <= set(week.columns)
    assert week["Stuks"].sum() == 450


def test_dashboard_zonder_stuks_ongewijzigd():
    holiday = pd.DataFrame(columns=["Datum", "Omschrijving", "Type"])
    inp = _dashboard_input().drop(columns=["Stuks", "Stuks_bron"])
    _, _, dag, week, _ = build_dashboard_data(inp, holiday, date(2026, 10, 12), 60000, 2, 2000)
    assert not any(c.startswith("Stuks") or c.endswith("_stuks") for c in dag.columns)
    assert "Stuks" not in week.columns
