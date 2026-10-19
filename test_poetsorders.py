"""
Tests voor de poetsorder-regel (v2.5.24): poetsorders verzinken X werkdagen
vóór de leverdatum (CAL: 3), depotorders blijven op 1. Standaard uit
([locatie] poetsen_verzink_werkdagen = 0), zodat Groningen ongewijzigd blijft.
"""
import sys
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))

from shared import (  # noqa: E402
    build_dashboard_data,
    get_poetsen_verzink_werkdagen,
    markeer_poetsorders,
)

DO = pd.Timestamp("2026-10-22")   # donderdag


def _rows() -> pd.DataFrame:
    return pd.DataFrame([
        # bevestigd, poetsen in prijscategorie
        {"Nummer": "100V1", "PrijsCategorie": "Verzinken 2D licht,Poetsen licht ( B )",
         "Gewicht_bron": "Export+", "Ordernummer_base": 100},
        # bevestigd, VC-nummer maar prijscategorie zonder poetsen → geen poetsorder
        {"Nummer": "101VC1-V", "PrijsCategorie": "Verzinken 3D",
         "Gewicht_bron": "Export+", "Ordernummer_base": 101},
        # bevestigd, VC-nummer en prijscategorie leeg → poetsorder (terugval)
        {"Nummer": "102VC2-V", "PrijsCategorie": np.nan,
         "Gewicht_bron": "OrderExport2G verdeeld", "Ordernummer_base": 102},
        # coat-deel → nooit gemarkeerd
        {"Nummer": "100VC1-C", "PrijsCategorie": "Poetsen zwaar ( A )",
         "Gewicht_bron": "Export+", "Ordernummer_base": 100},
        # gewone bevestigde order
        {"Nummer": "103V1", "PrijsCategorie": "Verzinken 1D onder 10/ton",
         "Gewicht_bron": "Export+", "Ordernummer_base": 103},
        # reserveringen
        {"Nummer": "200", "Gewicht_bron": "Reservering", "Ordernummer_base": 200, "Gewicht poetsen": 0.5},
        {"Nummer": "201", "Gewicht_bron": "Reservering", "Ordernummer_base": 201,
         "Leverdatum C": pd.Timestamp("2026-10-30")},
        {"Nummer": "202", "Gewicht_bron": "Reservering", "Ordernummer_base": 202},
        {"Nummer": "203", "Gewicht_bron": "Reservering", "Ordernummer_base": 203},
    ])


def _regels() -> pd.DataFrame:
    return pd.DataFrame({
        "OrderID": [202, 203],
        "Subbonnummer V": ["202VC1-V", "203V1"],
        "aantal": [5, 5],
    })


def test_default_uit():
    assert get_poetsen_verzink_werkdagen() == 0


def test_markeer_poetsorders():
    out = markeer_poetsorders(_rows(), _regels())
    assert list(out["Poetsorder"]) == ["Ja", "Nee", "Ja", "Nee", "Nee", "Ja", "Ja", "Ja", "Nee"]
    assert out.loc[0, "Poetsorder_reden"] == "PrijsCategorie bevat poetsen"
    assert out.loc[2, "Poetsorder_reden"] == "Nummer VCx-V (PrijsCategorie leeg)"
    assert out.loc[5, "Poetsorder_reden"] == "Gewicht poetsen > 0"
    assert out.loc[6, "Poetsorder_reden"] == "Leverdatum C gevuld"
    assert out.loc[7, "Poetsorder_reden"] == "Orderregel VCx-V"


def test_markeer_zonder_orderregels_en_kolommen():
    df = pd.DataFrame({"Nummer": ["1V1"], "Gewicht_bron": ["Export+"]})
    out = markeer_poetsorders(df, None)
    assert out.loc[0, "Poetsorder"] == "Nee"


def _dash_input() -> pd.DataFrame:
    return pd.DataFrame({
        "Nummer": ["1V1", "2V1", "3V1"],
        "Status": ["Gelost"] * 3,
        "Verzinkstatus": ["Niet verzinkt"] * 3,
        "Leverdatum": [DO, DO, DO],
        "Gewicht_effectief_kg": [1000.0, 1000.0, 1000.0],
        "Gewicht_bron": ["Export+"] * 3,
        "Aanleveren depot": [0, 0, 1],
        "Poetsorder": ["Ja", "Nee", "Ja"],
    })


def _vzd(poetsen_offset):
    hol = pd.DataFrame(columns=["Datum", "Omschrijving", "Type"])
    df, *_ = build_dashboard_data(_dash_input(), hol, date(2026, 10, 12), 60000, 2, 2000,
                                  poetsen_offset=poetsen_offset)
    return list(df["Verzinkdatum"])


def test_verzinkdatum_poetsorder_min_3_depot_blijft_1():
    assert _vzd(3) == [pd.Timestamp("2026-10-19"),   # poets: do − 3 werkdagen = ma
                       pd.Timestamp("2026-10-20"),   # normaal: do − 2 = di
                       pd.Timestamp("2026-10-21")]   # depot poets: do − 1 = wo


def test_verzinkdatum_uit_ongewijzigd():
    assert _vzd(0) == [pd.Timestamp("2026-10-20"), pd.Timestamp("2026-10-20"),
                       pd.Timestamp("2026-10-21")]
