"""Unit tests voor het productie-dashboard (v2.5.13).

Test focus:
  - build_productie_dashboard_week: correcte samenvoeging van MIS + plan + handmatig
  - Weekend/feestdag-uitsluiting (tenzij MIS-data aanwezig)
  - TONNAGE PLAN: alleen voor peildatum en verder in iteratie 1
  - Norm-getters: scalar en periode-fallback
"""
import io
import json
import sys
from datetime import date
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import pytest


sys.path.insert(0, str(Path(__file__).parent))

# Import na sys.path-tweak zodat pytest vanuit repo-root werkt
from shared import (  # noqa: E402
    MIS_KOLOMMEN,
    PRODUCTIE_DASHBOARD_KOLOMMEN,
    build_productie_dashboard_week,
    _resolve_norm_op_datum,
    load_dashboard_manual,
)


# ── Fixtures ──────────────────────────────────────────────────────────────────

def _mis_row(datum, kg=None, manuurton=None, afkeur=None, traversen=None, m2trav=None):
    return {
        "Datum": pd.Timestamp(datum),
        "KG": kg,
        "ManuurTon": manuurton,
        "KG_Afkeur": afkeur,
        "Traverses": traversen,
        "M2Trav": m2trav,
    }


def _dag_row(verzinkdatum, gewicht_kg):
    return {"Verzinkdatum": pd.Timestamp(verzinkdatum), "Gewicht_kg": float(gewicht_kg)}


# ── build_productie_dashboard_week ────────────────────────────────────────────

class TestBuildProductieDashboardWeek:
    """Controleer de KPI-tabelbouw per ISO-week."""

    def test_lege_input_geeft_7_lege_rijen(self):
        result = build_productie_dashboard_week(
            mis_df=pd.DataFrame(columns=MIS_KOLOMMEN),
            dag_df=pd.DataFrame(columns=["Verzinkdatum", "Gewicht_kg"]),
            manual_dict={},
            jaar=2026,
            weeknr=35,
        )
        assert len(result) == 7
        assert list(result.columns) == PRODUCTIE_DASHBOARD_KOLOMMEN
        assert list(result["Dag"]) == ["Ma", "Di", "Wo", "Do", "Vr", "Za", "Zo"]

    def test_week_35_2026_maandag_is_24_augustus(self):
        result = build_productie_dashboard_week(
            mis_df=pd.DataFrame(columns=MIS_KOLOMMEN),
            dag_df=pd.DataFrame(columns=["Verzinkdatum", "Gewicht_kg"]),
            manual_dict={},
            jaar=2026,
            weeknr=35,
        )
        # ISO-week 35 van 2026: ma = 24 aug
        assert result.iloc[0]["Datum"] == date(2026, 8, 24)
        assert result.iloc[6]["Datum"] == date(2026, 8, 30)

    def test_mis_data_wordt_correct_gekoppeld(self):
        mis_df = pd.DataFrame([
            _mis_row("2026-08-24", kg=80000, manuurton=5.5, afkeur=100, traversen=80, m2trav=1000),
        ])
        result = build_productie_dashboard_week(
            mis_df=mis_df,
            dag_df=pd.DataFrame(columns=["Verzinkdatum", "Gewicht_kg"]),
            manual_dict={},
            jaar=2026,
            weeknr=35,
        )
        maandag = result.iloc[0]
        assert maandag["TONNAGE WERKELIJK"] == 80000.0
        assert maandag["MANUREN / TON"] == 5.5
        assert maandag["AFKEUR IN KG"] == 100.0
        assert maandag["AANTAL TRAVERSEN"] == 80.0
        assert maandag["GEM GEWICHT PER TR"] == 1000.0
        # Dinsdag heeft geen MIS-data → alle MIS-velden leeg (None wordt NaN in float-kolom)
        dinsdag = result.iloc[1]
        assert pd.isna(dinsdag["TONNAGE WERKELIJK"])
        assert pd.isna(dinsdag["MANUREN / TON"])

    def test_weekend_zonder_data_krijgt_lege_kpi_velden(self):
        result = build_productie_dashboard_week(
            mis_df=pd.DataFrame(columns=MIS_KOLOMMEN),
            dag_df=pd.DataFrame(columns=["Verzinkdatum", "Gewicht_kg"]),
            manual_dict={},
            jaar=2026,
            weeknr=35,
        )
        zaterdag = result.iloc[5]
        zondag = result.iloc[6]
        for kpi in PRODUCTIE_DASHBOARD_KOLOMMEN[2:]:
            assert pd.isna(zaterdag[kpi]), f"Za {kpi} moet leeg zijn zonder MIS-data"
            assert pd.isna(zondag[kpi]), f"Zo {kpi} moet leeg zijn"

    def test_gewerkte_zaterdag_krijgt_normale_rij(self):
        # Zaterdag 29 aug 2026 met MIS-data
        mis_df = pd.DataFrame([
            _mis_row("2026-08-29", kg=45000, manuurton=7.2, afkeur=50, traversen=45, m2trav=1000),
        ])
        result = build_productie_dashboard_week(
            mis_df=mis_df,
            dag_df=pd.DataFrame(columns=["Verzinkdatum", "Gewicht_kg"]),
            manual_dict={},
            jaar=2026,
            weeknr=35,
        )
        zaterdag = result.iloc[5]
        assert zaterdag["TONNAGE WERKELIJK"] == 45000.0
        assert zaterdag["AANTAL TRAVERSEN"] == 45.0

    def test_feestdag_zonder_mis_data_blijft_leeg(self):
        # 27 apr 2026 (Koningsdag, ma) → week 18
        feestdagen = {pd.Timestamp("2026-04-27")}
        result = build_productie_dashboard_week(
            mis_df=pd.DataFrame(columns=MIS_KOLOMMEN),
            dag_df=pd.DataFrame(columns=["Verzinkdatum", "Gewicht_kg"]),
            manual_dict={},
            jaar=2026,
            weeknr=18,
            feestdagen=feestdagen,
        )
        maandag = result.iloc[0]
        assert maandag["Datum"] == date(2026, 4, 27)
        assert pd.isna(maandag["TONNAGE WERKELIJK"])
        assert pd.isna(maandag["MANUREN / TON"])

    def test_handmatige_invoer_wordt_opgenomen(self):
        manual = {
            "2026-08-24": {"klachten": 2, "storingstijd_min": 45, "veiligheidsincidenten": 0},
        }
        result = build_productie_dashboard_week(
            mis_df=pd.DataFrame(columns=MIS_KOLOMMEN),
            dag_df=pd.DataFrame(columns=["Verzinkdatum", "Gewicht_kg"]),
            manual_dict=manual,
            jaar=2026,
            weeknr=35,
        )
        maandag = result.iloc[0]
        assert maandag["AANTAL KLACHTEN"] == 2
        assert maandag["Storingstijd in min"] == 45
        assert maandag["VEILIGHEIDSINCIDENTEN"] == 0
        # Dinsdag zonder invoer blijft leeg (None → NaN in numerieke kolom)
        assert pd.isna(result.iloc[1]["AANTAL KLACHTEN"])

    def test_tonnage_plan_alleen_voor_vandaag_en_verder(self):
        # dag_df bevat zowel een historische als toekomstige datum
        today = pd.Timestamp(date.today()).normalize()
        gisteren = today - pd.Timedelta(days=1)
        morgen = today + pd.Timedelta(days=1)
        dag_df = pd.DataFrame([
            _dag_row(gisteren, 55000),
            _dag_row(today, 60000),
            _dag_row(morgen, 62000),
        ])
        # Gebruik de ISO-week die vandaag omvat
        iso = today.isocalendar()
        result = build_productie_dashboard_week(
            mis_df=pd.DataFrame(columns=MIS_KOLOMMEN),
            dag_df=dag_df,
            manual_dict={},
            jaar=iso.year,
            weeknr=iso.week,
        )
        for _, r in result.iterrows():
            d = pd.Timestamp(r["Datum"])
            if d < today:
                assert pd.isna(r["TONNAGE PLAN"]), f"{d.date()} zou historisch geen plan mogen hebben"
            elif d == today:
                assert r["TONNAGE PLAN"] == 60000.0
            elif d == morgen:
                assert r["TONNAGE PLAN"] == 62000.0


# ── Norm-getters ──────────────────────────────────────────────────────────────

class TestNormResolver:
    """Controleer de fallback-volgorde: periode → scalar → default."""

    def test_default_gebruikt_als_geen_secrets(self):
        # Wanneer st.secrets['locatie'] geen key heeft, gebruikt hij de default
        with patch("shared._get_locatie_list", return_value=[]), \
             patch("shared._get_locatie_number", return_value=6.0):
            assert _resolve_norm_op_datum(date(2026, 8, 27), "x", "y", 6.0) == 6.0

    def test_scalar_uit_secrets_wint_van_default(self):
        with patch("shared._get_locatie_list", return_value=[]), \
             patch("shared._get_locatie_number", return_value=5.5):
            assert _resolve_norm_op_datum(date(2026, 8, 27), "x", "y", 6.0) == 5.5

    def test_periode_wint_van_scalar(self):
        periods = [
            {"van": "2026-01-01", "tot": "2026-06-30", "waarde": 6.0},
            {"van": "2026-07-01", "tot": "9999-12-31", "waarde": 5.0},
        ]
        with patch("shared._get_locatie_list", return_value=periods), \
             patch("shared._get_locatie_number", return_value=99.9):
            # Datum in eerste periode → 6.0
            assert _resolve_norm_op_datum(date(2026, 3, 15), "x", "y", 8.0) == 6.0
            # Datum in tweede periode → 5.0
            assert _resolve_norm_op_datum(date(2026, 9, 1), "x", "y", 8.0) == 5.0

    def test_geen_matchende_periode_valt_terug_op_scalar(self):
        periods = [{"van": "2026-01-01", "tot": "2026-06-30", "waarde": 6.0}]
        with patch("shared._get_locatie_list", return_value=periods), \
             patch("shared._get_locatie_number", return_value=5.5):
            assert _resolve_norm_op_datum(date(2026, 9, 1), "x", "y", 8.0) == 5.5

    def test_periode_zonder_tot_is_open_einde(self):
        periods = [{"van": "2026-07-01", "waarde": 5.5}]
        with patch("shared._get_locatie_list", return_value=periods), \
             patch("shared._get_locatie_number", return_value=99.9):
            assert _resolve_norm_op_datum(date(2099, 1, 1), "x", "y", 8.0) == 5.5


# ── Load handmatig — lege fallback ────────────────────────────────────────────

class TestLoadDashboardManual:
    def test_missing_file_returns_empty_dict(self):
        with patch("shared.download_file", side_effect=Exception("not found")):
            assert load_dashboard_manual() == {}

    def test_corrupt_json_returns_empty_dict(self):
        with patch("shared.download_file", return_value=b"not json"):
            assert load_dashboard_manual() == {}

    def test_niet_dict_returns_empty_dict(self):
        with patch("shared.download_file", return_value=b'["a","b"]'):
            assert load_dashboard_manual() == {}

    def test_geldige_json_wordt_teruggegeven(self):
        payload = {"2026-08-27": {"klachten": 1, "storingstijd_min": 10, "veiligheidsincidenten": 0}}
        with patch("shared.download_file", return_value=json.dumps(payload).encode("utf-8")):
            assert load_dashboard_manual() == payload
