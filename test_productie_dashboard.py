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
    SNAPSHOT_RETENTION_DAYS,
    build_productie_dashboard_week,
    build_productie_dashboard_ytd,
    build_otif_trend_df,
    _resolve_norm_op_datum,
    load_dashboard_manual,
    load_mis_data,
    load_daily_snapshots,
    save_daily_snapshot,
)


@pytest.fixture(autouse=True)
def _clear_streamlit_cache():
    """Zorg dat elke test met een leeg cache begint. Nodig omdat de load- en
    build-functies met @st.cache_data zijn gedecoreerd; zonder clear zouden
    mock-patches in opvolgende tests een cache-hit teruggeven."""
    for fn in (load_dashboard_manual, load_mis_data, load_daily_snapshots,
               build_productie_dashboard_week, build_productie_dashboard_ytd):
        try:
            fn.clear()
        except AttributeError:
            pass
    yield


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


# ── build_productie_dashboard_ytd ─────────────────────────────────────────────

class TestBuildProductieDashboardYtd:
    """Controleer de YTD-aggregatie per ISO-week."""

    def test_lege_input_geeft_leeg_frame(self):
        result = build_productie_dashboard_ytd(pd.DataFrame(columns=MIS_KOLOMMEN), jaar=2026)
        assert result.empty
        assert list(result.columns) == [
            "Weeknr", "Week_startdatum", "Werkelijk_kg_totaal",
            "Manuren_per_ton_gewogen", "Traversen_totaal", "Gem_gewicht_per_traverse",
        ]

    def test_aggregatie_per_week(self):
        # Week 35 (24-30 aug 2026): 3 dagen met data
        mis_df = pd.DataFrame([
            _mis_row("2026-08-24", kg=70000, manuurton=6.0, afkeur=100, traversen=70, m2trav=1000),
            _mis_row("2026-08-25", kg=80000, manuurton=5.5, afkeur=200, traversen=80, m2trav=1000),
            _mis_row("2026-08-26", kg=50000, manuurton=7.0, afkeur=50,  traversen=50, m2trav=1000),
        ])
        result = build_productie_dashboard_ytd(mis_df, jaar=2026)
        # Alle drie in dezelfde week
        assert len(result) == 1
        row = result.iloc[0]
        assert row["Weeknr"] == 35
        assert row["Werkelijk_kg_totaal"] == 200000
        assert row["Traversen_totaal"] == 200
        # Rekenkundig gemiddelde: (6.0 + 5.5 + 7.0) / 3 = 6.1667
        assert abs(row["Manuren_per_ton_gewogen"] - (6.0 + 5.5 + 7.0) / 3) < 0.001
        # Gem gewicht: 200000 / 200 = 1000
        assert row["Gem_gewicht_per_traverse"] == 1000.0

    def test_alleen_gevraagd_jaar(self):
        # Let op: december-datums vanaf ~28e vallen ISO-technisch al in week 1
        # van het volgende jaar. 15 dec zit gegarandeerd in ISO-jaar 2025.
        mis_df = pd.DataFrame([
            _mis_row("2025-12-15", kg=50000, manuurton=6.0, traversen=50),
            _mis_row("2026-01-05", kg=60000, manuurton=6.0, traversen=60),
        ])
        result_2026 = build_productie_dashboard_ytd(mis_df, jaar=2026)
        assert len(result_2026) == 1
        assert result_2026.iloc[0]["Werkelijk_kg_totaal"] == 60000

    def test_toekomstige_datums_worden_genegeerd(self):
        # Datum in de toekomst
        toekomst = pd.Timestamp(date.today()) + pd.Timedelta(days=30)
        mis_df = pd.DataFrame([
            _mis_row(toekomst, kg=10000, manuurton=6.0, traversen=10),
        ])
        result = build_productie_dashboard_ytd(mis_df, jaar=toekomst.year)
        assert result.empty

    def test_rijen_zonder_kg_worden_overgeslagen(self):
        mis_df = pd.DataFrame([
            _mis_row("2026-08-24", kg=None, manuurton=6.0, traversen=None),  # skip
            _mis_row("2026-08-25", kg=80000, manuurton=5.5, traversen=80, m2trav=1000),
        ])
        result = build_productie_dashboard_ytd(mis_df, jaar=2026)
        assert len(result) == 1
        assert result.iloc[0]["Werkelijk_kg_totaal"] == 80000

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


# ── Daily snapshots ───────────────────────────────────────────────────────────

class TestSaveDailySnapshot:
    """Controleer save-gedrag: overwrite, backfilled-vlag, retentie."""

    def test_save_nieuwe_datum_retourneert_true(self):
        with patch("shared.download_file", side_effect=Exception("not found")), \
             patch("shared.upload_file") as m_upload:
            otif = {"otif_pct": 85.0, "totaal": 40, "gereed": 34, "niet_gereed": 6}
            saved = save_daily_snapshot(date(2026, 8, 25), 55000, otif)
            assert saved is True
            assert m_upload.called
            # De payload naar upload moet de nieuwe entry bevatten
            _fname, data = m_upload.call_args[0]
            body = json.loads(data.decode("utf-8"))
            assert "2026-08-25" in body
            entry = body["2026-08-25"]
            assert entry["tonnage_plan_kg"] == 55000.0
            assert entry["otif_pct"] == 85.0
            assert entry["otif_totaal"] == 40
            assert entry["backfilled"] is False

    def test_save_bestaande_datum_zonder_overwrite_retourneert_false(self):
        existing = {"2026-08-25": {"tonnage_plan_kg": 55000, "otif_pct": 85.0,
                                     "otif_totaal": 40, "otif_gereed": 34,
                                     "otif_niet_gereed": 6, "created_at": "2026-08-26T08:00:00",
                                     "backfilled": False}}
        with patch("shared.download_file", return_value=json.dumps(existing).encode("utf-8")), \
             patch("shared.upload_file") as m_upload:
            saved = save_daily_snapshot(date(2026, 8, 25), 60000, {"otif_pct": 90}, overwrite=False)
            assert saved is False
            assert not m_upload.called

    def test_save_met_overwrite_vervangt_bestaande_entry(self):
        existing = {"2026-08-25": {"tonnage_plan_kg": 55000, "otif_pct": 85.0,
                                     "otif_totaal": 40, "otif_gereed": 34,
                                     "otif_niet_gereed": 6, "created_at": "2026-08-26T08:00:00",
                                     "backfilled": False}}
        with patch("shared.download_file", return_value=json.dumps(existing).encode("utf-8")), \
             patch("shared.upload_file") as m_upload:
            saved = save_daily_snapshot(date(2026, 8, 25), 60000,
                                          {"otif_pct": 90, "totaal": 45, "gereed": 41, "niet_gereed": 4},
                                          overwrite=True, backfilled=True)
            assert saved is True
            body = json.loads(m_upload.call_args[0][1].decode("utf-8"))
            entry = body["2026-08-25"]
            assert entry["tonnage_plan_kg"] == 60000.0
            assert entry["otif_pct"] == 90.0
            assert entry["backfilled"] is True

    def test_retentie_verwijdert_oude_entries(self):
        oud = (pd.Timestamp(date.today()) - pd.Timedelta(days=SNAPSHOT_RETENTION_DAYS + 10)).strftime("%Y-%m-%d")
        recent = (pd.Timestamp(date.today()) - pd.Timedelta(days=5)).strftime("%Y-%m-%d")
        existing = {
            oud:    {"tonnage_plan_kg": 1, "otif_pct": 50, "otif_totaal": 1, "otif_gereed": 0,
                     "otif_niet_gereed": 1, "created_at": "old", "backfilled": False},
            recent: {"tonnage_plan_kg": 2, "otif_pct": 80, "otif_totaal": 2, "otif_gereed": 1,
                     "otif_niet_gereed": 1, "created_at": "recent", "backfilled": False},
        }
        with patch("shared.download_file", return_value=json.dumps(existing).encode("utf-8")), \
             patch("shared.upload_file") as m_upload:
            save_daily_snapshot(date.today(), 5000, {"otif_pct": 90})
            body = json.loads(m_upload.call_args[0][1].decode("utf-8"))
            assert oud not in body, "Oude entry buiten retentie zou weg moeten"
            assert recent in body, "Recente entry moet blijven"
            assert date.today().strftime("%Y-%m-%d") in body


class TestBuildOtifTrendDf:
    """Controleer de trend-DataFrame voor de grafiek."""

    def test_leeg_snapshots_geeft_leeg_frame(self):
        result = build_otif_trend_df({})
        assert result.empty
        assert list(result.columns) == ["Datum", "OTIF_pct", "Aantal_orders", "Backfilled"]

    def test_entries_zonder_otif_pct_worden_overgeslagen(self):
        snaps = {
            "2026-08-25": {"tonnage_plan_kg": 50000, "otif_pct": None, "otif_totaal": 0,
                           "otif_gereed": 0, "otif_niet_gereed": 0, "created_at": "x", "backfilled": False},
            "2026-08-26": {"tonnage_plan_kg": 60000, "otif_pct": 85.0, "otif_totaal": 40,
                           "otif_gereed": 34, "otif_niet_gereed": 6, "created_at": "x", "backfilled": False},
        }
        result = build_otif_trend_df(snaps, days_back=365)
        assert len(result) == 1
        assert result.iloc[0]["OTIF_pct"] == 85.0

    def test_gesorteerd_op_datum_en_gefilterd_op_days_back(self):
        vandaag = pd.Timestamp(date.today()).normalize()
        oud = vandaag - pd.Timedelta(days=200)
        recent = vandaag - pd.Timedelta(days=10)
        snaps = {
            oud.strftime("%Y-%m-%d"):    {"tonnage_plan_kg": 1, "otif_pct": 70, "otif_totaal": 10,
                                            "otif_gereed": 7, "otif_niet_gereed": 3, "created_at": "x", "backfilled": False},
            recent.strftime("%Y-%m-%d"): {"tonnage_plan_kg": 2, "otif_pct": 85, "otif_totaal": 20,
                                            "otif_gereed": 17, "otif_niet_gereed": 3, "created_at": "x", "backfilled": True},
        }
        # 90 dagen: alleen recent
        result = build_otif_trend_df(snaps, days_back=90)
        assert len(result) == 1
        assert result.iloc[0]["OTIF_pct"] == 85
        assert bool(result.iloc[0]["Backfilled"]) is True
        # 365 dagen: beide, oplopend gesorteerd
        result_full = build_otif_trend_df(snaps, days_back=365)
        assert len(result_full) == 2
        assert list(result_full["OTIF_pct"]) == [70.0, 85.0]


class TestBuildProductieDashboardWeekWithSnapshots:
    """Controleer dat historische TONNAGE PLAN uit snapshots komt."""

    def test_historisch_uit_snapshot(self):
        # Kies een dag in het (recente) verleden
        gisteren = pd.Timestamp(date.today()) - pd.Timedelta(days=1)
        iso = gisteren.isocalendar()
        snap_key = gisteren.strftime("%Y-%m-%d")
        snapshots = {
            snap_key: {"tonnage_plan_kg": 62000, "otif_pct": 88.0, "otif_totaal": 40,
                       "otif_gereed": 35, "otif_niet_gereed": 5, "created_at": "x", "backfilled": False}
        }
        result = build_productie_dashboard_week(
            mis_df=pd.DataFrame(columns=MIS_KOLOMMEN),
            dag_df=pd.DataFrame(columns=["Verzinkdatum", "Gewicht_kg"]),
            manual_dict={},
            jaar=iso.year,
            weeknr=iso.week,
            feestdagen=frozenset(),
            snapshots=snapshots,
        )
        # Zoek de rij voor gisteren
        hit = result[pd.to_datetime(result["Datum"]) == gisteren.normalize()]
        assert len(hit) == 1
        assert hit.iloc[0]["TONNAGE PLAN"] == 62000.0

    def test_historisch_zonder_snapshot_blijft_leeg(self):
        gisteren = pd.Timestamp(date.today()) - pd.Timedelta(days=1)
        iso = gisteren.isocalendar()
        result = build_productie_dashboard_week(
            mis_df=pd.DataFrame(columns=MIS_KOLOMMEN),
            dag_df=pd.DataFrame(columns=["Verzinkdatum", "Gewicht_kg"]),
            manual_dict={},
            jaar=iso.year,
            weeknr=iso.week,
            feestdagen=frozenset(),
            snapshots={},
        )
        hit = result[pd.to_datetime(result["Datum"]) == gisteren.normalize()]
        assert pd.isna(hit.iloc[0]["TONNAGE PLAN"])

    def test_toekomst_gebruikt_nog_altijd_dag_df(self):
        morgen = pd.Timestamp(date.today()) + pd.Timedelta(days=1)
        iso = morgen.isocalendar()
        dag_df = pd.DataFrame([{"Verzinkdatum": morgen, "Gewicht_kg": 58000.0}])
        result = build_productie_dashboard_week(
            mis_df=pd.DataFrame(columns=MIS_KOLOMMEN),
            dag_df=dag_df,
            manual_dict={},
            jaar=iso.year,
            weeknr=iso.week,
            feestdagen=frozenset(),
            snapshots={},
        )
        hit = result[pd.to_datetime(result["Datum"]) == morgen.normalize()]
        assert hit.iloc[0]["TONNAGE PLAN"] == 58000.0
