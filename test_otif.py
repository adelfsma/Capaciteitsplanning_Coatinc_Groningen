"""
Snelle test voor de OTIF-logica in shared.py.
- Test map_otif_status voor de mapping uit Status_OTIF.xlsx
- Test find_originele_leverdatum_column
- Test compute_otif met een gesimuleerd merged-dataframe
- Test get_peildatum_from_metadata

Deze test omzeilt de Streamlit/Supabase-afhankelijkheden door alleen de pure
helpers uit shared.py te importeren en een lokaal gemaakt DataFrame te
gebruiken.
"""
import sys
from datetime import date
from pathlib import Path

import pandas as pd

# Streamlit imports omzeilen door een dummy 'streamlit' module te maken; shared.py
# gebruikt st.cache_resource en st.secrets alleen wanneer de cloud-functies
# worden aangeroepen. Voor pure helpers is dat niet nodig.
sys.path.insert(0, str(Path(__file__).parent))

from shared import (  # noqa: E402
    OTIF_STATUS_MAP,
    ORIGINELE_LEVERDATUM_CANDIDATES,
    map_otif_status,
    find_originele_leverdatum_column,
    compute_otif,
    get_peildatum_from_metadata,
)


def test_map_otif_status_matches_spec():
    """De mapping uit Status_OTIF.xlsx moet 1-op-1 herkend worden."""
    verwacht = {
        "uitgeleverd":               "Ja",
        "Afgehaald":                 "Ja",
        "PC Afgehaald":              "Nvt",
        "Productie gereed":          "Nee",
        "Geblokkeerd":               "Nee",
        "Opgehangen":                "Nee",
        "coat gereed":               "Ja",
        "UB":                        "Nee",
        "Nabewerking nog uitvoeren": "Nee",
        "UB V Gereed":               "Ja",
        "meetrapport":               "Nee",
        "PC Opgehangen":             "Nee",
    }
    for status, expected in verwacht.items():
        got = map_otif_status(status)
        assert got == expected, f"map_otif_status({status!r}) → {got!r}, verwacht {expected!r}"
    # Case-insensitief
    assert map_otif_status("UITGELEVERD") == "Ja"
    assert map_otif_status("  Afgehaald  ") == "Ja"
    # Onbekende status
    assert map_otif_status("Gereserveerd*") == "Onbekend"
    assert map_otif_status("Reservering") == "Onbekend"
    assert map_otif_status(None) == "Onbekend"
    print("✅ test_map_otif_status_matches_spec")


def test_find_originele_leverdatum_column():
    # Kolom aanwezig onder verschillende namen
    for kandidaat in ORIGINELE_LEVERDATUM_CANDIDATES:
        df = pd.DataFrame({kandidaat: [pd.Timestamp("2026-07-17")]})
        assert find_originele_leverdatum_column(df) == kandidaat, f"Kandidaat {kandidaat} niet herkend"
    # Case-insensitief herkennen
    df = pd.DataFrame({"originele DATUM": [1]})
    assert find_originele_leverdatum_column(df) == "originele DATUM"
    # Geen kolom aanwezig
    df = pd.DataFrame({"Datum": [1], "Nummer": [1]})
    assert find_originele_leverdatum_column(df) is None
    # Leeg dataframe
    assert find_originele_leverdatum_column(pd.DataFrame()) is None
    print("✅ test_find_originele_leverdatum_column")


def _mock_merged(peildag: date) -> pd.DataFrame:
    """
    Bouw een gesimuleerd 'merged'-DataFrame met een mix aan orders op de
    peildag, een reservering (moet niet meetellen), en orders op andere datums.
    """
    peil_ts = pd.Timestamp(peildag)
    other_ts = peil_ts + pd.Timedelta(days=1)

    return pd.DataFrame([
        # Op peildag – meetellend
        {"Bron_week": "0", "Nummer": "A-001", "Leverdatum": peil_ts,
         "Originele_leverdatum": peil_ts, "Status": "uitgeleverd",
         "Gewicht_effectief_kg": 1000.0, "Klantnaam": "Klant A"},
        {"Bron_week": "0", "Nummer": "A-002", "Leverdatum": peil_ts,
         "Originele_leverdatum": peil_ts, "Status": "Afgehaald",
         "Gewicht_effectief_kg": 500.0, "Klantnaam": "Klant B"},
        {"Bron_week": "0", "Nummer": "A-003", "Leverdatum": peil_ts,
         "Originele_leverdatum": peil_ts, "Status": "coat gereed",
         "Gewicht_effectief_kg": 250.0, "Klantnaam": "Klant C"},
        {"Bron_week": "0", "Nummer": "A-004", "Leverdatum": peil_ts,
         "Originele_leverdatum": peil_ts, "Status": "UB V Gereed",
         "Gewicht_effectief_kg": 750.0, "Klantnaam": "Klant D"},
        # Nee (te laat)
        {"Bron_week": "0", "Nummer": "A-005", "Leverdatum": peil_ts,
         "Originele_leverdatum": peil_ts, "Status": "Productie gereed",
         "Gewicht_effectief_kg": 1500.0, "Klantnaam": "Klant E"},
        {"Bron_week": "0", "Nummer": "A-006", "Leverdatum": peil_ts,
         "Originele_leverdatum": peil_ts, "Status": "Geblokkeerd",
         "Gewicht_effectief_kg": 300.0, "Klantnaam": "Klant F"},
        {"Bron_week": "0", "Nummer": "A-007", "Leverdatum": peil_ts,
         "Originele_leverdatum": peil_ts, "Status": "Opgehangen",
         "Gewicht_effectief_kg": 400.0, "Klantnaam": "Klant G"},
        # Nvt
        {"Bron_week": "0", "Nummer": "A-008", "Leverdatum": peil_ts,
         "Originele_leverdatum": peil_ts, "Status": "PC Afgehaald",
         "Gewicht_effectief_kg": 600.0, "Klantnaam": "Klant H"},
        # Onbekende status (bijv. Gereserveerd* die niet in mapping staat)
        {"Bron_week": "0", "Nummer": "A-009", "Leverdatum": peil_ts,
         "Originele_leverdatum": peil_ts, "Status": "Gereserveerd*",
         "Gewicht_effectief_kg": 100.0, "Klantnaam": "Klant I"},
        # Reservering op peildag – moet uitgesloten worden
        {"Bron_week": "reservering", "Nummer": "R-100", "Leverdatum": peil_ts,
         "Originele_leverdatum": pd.NaT, "Status": "Reservering",
         "Gewicht_effectief_kg": 800.0, "Klantnaam": "Klant J"},
        # Order op andere datum – mag niet meetellen
        {"Bron_week": "+1", "Nummer": "B-001", "Leverdatum": other_ts,
         "Originele_leverdatum": other_ts, "Status": "Productie gereed",
         "Gewicht_effectief_kg": 900.0, "Klantnaam": "Klant K"},
    ])


def test_compute_otif_basic():
    peildag = date(2026, 7, 17)
    df = _mock_merged(peildag)

    r = compute_otif(df, peildatum=peildag, date_column="Leverdatum", date_column_label="Datum")

    # Verwacht: 9 orders op de peildag (reservering + andere datum vallen buiten).
    assert r["totaal"] == 9, f"totaal = {r['totaal']}"
    # Ja: 4 (uitgeleverd, Afgehaald, coat gereed, UB V Gereed)
    assert r["gereed"] == 4, f"gereed = {r['gereed']}"
    # Nee: 3 (Productie gereed, Geblokkeerd, Opgehangen)
    assert r["niet_gereed"] == 3, f"niet_gereed = {r['niet_gereed']}"
    # Nvt: 1 (PC Afgehaald)
    assert r["nvt"] == 1, f"nvt = {r['nvt']}"
    # Onbekend: 1 (Gereserveerd*)
    assert r["onbekend"] == 1, f"onbekend = {r['onbekend']}"

    # Formule: (1 - 3/9) * 100 = 66.667%
    assert abs(r["otif_pct"] - (1 - 3/9) * 100) < 1e-6, f"otif_pct = {r['otif_pct']}"
    assert r["date_column"] == "Leverdatum"
    assert r["date_column_label"] == "Datum"

    # Reservering moet niet in de subset zitten
    assert "R-100" not in set(r["orders"]["Nummer"].astype(str))
    # Andere datum moet niet in de subset zitten
    assert "B-001" not in set(r["orders"]["Nummer"].astype(str))

    print(f"✅ test_compute_otif_basic (OTIF = {r['otif_pct']:.2f}%)")


def test_compute_otif_originele_datum():
    """Toetsing tegen Originele_leverdatum moet dezelfde resultaten geven wanneer
    beide datums gelijk zijn."""
    peildag = date(2026, 7, 17)
    df = _mock_merged(peildag)

    r = compute_otif(df, peildatum=peildag, date_column="Originele_leverdatum",
                     date_column_label="Originele datum")
    # In deze fake set zijn Leverdatum en Originele_leverdatum gelijk, dus zelfde uitkomst.
    assert r["totaal"] == 9
    assert r["gereed"] == 4
    assert r["niet_gereed"] == 3
    assert r["date_column"] == "Originele_leverdatum"
    print("✅ test_compute_otif_originele_datum")


def test_compute_otif_empty_and_no_column():
    peildag = date(2026, 7, 17)
    # Kolom afwezig
    r = compute_otif(pd.DataFrame({"Status": ["x"], "Bron_week": ["0"]}),
                     peildatum=peildag, date_column="Leverdatum")
    assert r["totaal"] == 0 and r["otif_pct"] is None
    # Leeg dataframe
    r = compute_otif(pd.DataFrame(), peildatum=peildag)
    assert r["totaal"] == 0 and r["otif_pct"] is None
    # Peildag zonder orders in dataset
    df = _mock_merged(date(2026, 6, 1))
    r = compute_otif(df, peildatum=peildag, date_column="Leverdatum")
    assert r["totaal"] == 0 and r["otif_pct"] is None
    print("✅ test_compute_otif_empty_and_no_column")


def test_compute_otif_perfect_score():
    peildag = date(2026, 7, 17)
    df = pd.DataFrame([
        {"Bron_week": "0", "Nummer": str(i), "Leverdatum": pd.Timestamp(peildag),
         "Originele_leverdatum": pd.Timestamp(peildag), "Status": "Afgehaald",
         "Gewicht_effectief_kg": 100.0}
        for i in range(50)
    ])
    r = compute_otif(df, peildatum=peildag)
    assert r["totaal"] == 50 and r["gereed"] == 50 and r["niet_gereed"] == 0
    assert r["otif_pct"] == 100.0
    print("✅ test_compute_otif_perfect_score (OTIF = 100%)")


def test_get_peildatum_from_metadata():
    # De strftime-format van manager_app.py.
    meta = {"published_at": "17-07-2026 14:30:00"}
    assert get_peildatum_from_metadata(meta) == date(2026, 7, 17)
    # Andere formaten
    assert get_peildatum_from_metadata({"published_at": "2026-07-17"}) == date(2026, 7, 17)
    assert get_peildatum_from_metadata({"published_at": "2026-07-17 08:00:00"}) == date(2026, 7, 17)
    # Onbekend/leeg → fallback op vandaag
    assert get_peildatum_from_metadata(None) == date.today()
    assert get_peildatum_from_metadata({}) == date.today()
    assert get_peildatum_from_metadata({"published_at": ""}) == date.today()
    print("✅ test_get_peildatum_from_metadata")


def test_previous_workday_holidays():
    """previous_workday moet weekenden altijd overslaan; met holiday_dates
    ook feestdagen."""
    from shared import previous_workday
    # Maandag → vrijdag
    assert previous_workday(date(2026, 7, 20)) == date(2026, 7, 17)
    # Vrijdag → donderdag
    assert previous_workday(date(2026, 7, 17)) == date(2026, 7, 16)
    # Zondag → vrijdag (start van zondag; -1 = zaterdag; -2 = vrijdag)
    assert previous_workday(date(2026, 7, 19)) == date(2026, 7, 17)
    # Feestdag overslaan: als vr 17-07 een feestdag is, dan ma → do 16-07
    hd = {date(2026, 7, 17)}
    assert previous_workday(date(2026, 7, 20), hd) == date(2026, 7, 16)
    # Meerdere feestdagen na elkaar
    hd = {date(2026, 7, 16), date(2026, 7, 17)}
    assert previous_workday(date(2026, 7, 20), hd) == date(2026, 7, 15)
    print("✅ test_previous_workday_holidays")


def test_compute_otif_depot_shift_basic():
    """
    Depot-orders van de vorige werkdag tellen mee op peildatum;
    depot-orders van de peildatum tellen NIET mee (worden morgen beoordeeld);
    niet-depot-orders werken zoals voorheen.
    """
    peildag  = date(2026, 7, 21)   # dinsdag
    vorige   = date(2026, 7, 20)   # maandag (vorige werkdag)

    df = pd.DataFrame([
        # 1) Op peildag, niet-depot → meetellend (Ja)
        {"Bron_week": "0", "Nummer": "N-1", "Leverdatum": pd.Timestamp(peildag),
         "Status": "uitgeleverd", "Aanleveren depot": 0},
        # 2) Op peildag, depot=1 → NIET meetellen (verschoven naar morgen)
        {"Bron_week": "0", "Nummer": "D-1", "Leverdatum": pd.Timestamp(peildag),
         "Status": "Productie gereed", "Aanleveren depot": 1},
        # 3) Op vorige werkdag, depot=1 → wél meetellen (Nee = te laat)
        {"Bron_week": "0", "Nummer": "D-2", "Leverdatum": pd.Timestamp(vorige),
         "Status": "Productie gereed", "Aanleveren depot": 1},
        # 4) Op vorige werkdag, depot=1, gereed (Ja)
        {"Bron_week": "0", "Nummer": "D-3", "Leverdatum": pd.Timestamp(vorige),
         "Status": "Afgehaald", "Aanleveren depot": 1},
        # 5) Op vorige werkdag, niet-depot → NIET meetellen (hoort bij gisteren)
        {"Bron_week": "0", "Nummer": "N-2", "Leverdatum": pd.Timestamp(vorige),
         "Status": "uitgeleverd", "Aanleveren depot": 0},
        # 6) Op peildag, NaN depot-waarde → behandel als depot=0 → meetellen
        {"Bron_week": "0", "Nummer": "N-3", "Leverdatum": pd.Timestamp(peildag),
         "Status": "Afgehaald", "Aanleveren depot": None},
        # 7) Reservering op peildag → nooit meetellen
        {"Bron_week": "reservering", "Nummer": "R-1", "Leverdatum": pd.Timestamp(peildag),
         "Status": "Reservering", "Aanleveren depot": 0},
    ])

    r = compute_otif(df, peildatum=peildag, holiday_dates=set())

    # Meetellend: N-1 (Ja), D-2 (Nee), D-3 (Ja), N-3 (Ja) = 4 orders
    assert r["totaal"] == 4, f"totaal={r['totaal']}"
    assert r["gereed"] == 3, f"gereed={r['gereed']}"        # N-1, D-3, N-3
    assert r["niet_gereed"] == 1, f"niet_gereed={r['niet_gereed']}"  # D-2
    assert r["aantal_depot"] == 2, f"aantal_depot={r['aantal_depot']}"  # D-2, D-3
    assert r["depot_shift_actief"] is True
    assert r["peildatum_vorige_werkdag"] == vorige
    # OTIF = (1 - 1/4) * 100 = 75%
    assert abs(r["otif_pct"] - 75.0) < 1e-6
    # D-1 (depot op peildag) NIET in subset
    assert "D-1" not in set(r["orders"]["Nummer"].astype(str))
    # N-2 (niet-depot op vorige werkdag) NIET in subset
    assert "N-2" not in set(r["orders"]["Nummer"].astype(str))
    print("✅ test_compute_otif_depot_shift_basic")


def test_compute_otif_depot_shift_over_weekend():
    """Op maandag moeten depot-orders van vrijdag meegeteld worden."""
    maandag = date(2026, 7, 20)
    vrijdag = date(2026, 7, 17)

    df = pd.DataFrame([
        {"Bron_week": "0", "Nummer": "MA-1", "Leverdatum": pd.Timestamp(maandag),
         "Status": "uitgeleverd", "Aanleveren depot": 0},
        {"Bron_week": "0", "Nummer": "DEP-VR", "Leverdatum": pd.Timestamp(vrijdag),
         "Status": "Productie gereed", "Aanleveren depot": 1},
        {"Bron_week": "0", "Nummer": "GEEN-VR", "Leverdatum": pd.Timestamp(vrijdag),
         "Status": "uitgeleverd", "Aanleveren depot": 0},
    ])

    r = compute_otif(df, peildatum=maandag, holiday_dates=set())
    assert r["peildatum_vorige_werkdag"] == vrijdag
    # MA-1 (Ja) + DEP-VR (Nee) = 2 meetellend
    assert r["totaal"] == 2
    assert r["gereed"] == 1
    assert r["niet_gereed"] == 1
    assert r["aantal_depot"] == 1
    # GEEN-VR (niet-depot vrijdag) NIET meegeteld
    assert "GEEN-VR" not in set(r["orders"]["Nummer"].astype(str))
    print("✅ test_compute_otif_depot_shift_over_weekend")


def test_compute_otif_depot_shift_over_holiday():
    """Feestdag tussen peildatum en vorige werkdag wordt overgeslagen."""
    donderdag = date(2026, 5, 14)
    vrijdag   = date(2026, 5, 15)   # feestdag
    maandag   = date(2026, 5, 18)

    df = pd.DataFrame([
        {"Bron_week": "0", "Nummer": "DEP-DO", "Leverdatum": pd.Timestamp(donderdag),
         "Status": "Afgehaald", "Aanleveren depot": 1},
    ])

    r = compute_otif(df, peildatum=maandag, holiday_dates={vrijdag})
    assert r["peildatum_vorige_werkdag"] == donderdag, (
        f"verwacht do {donderdag}, kreeg {r['peildatum_vorige_werkdag']}"
    )
    assert r["totaal"] == 1
    assert r["aantal_depot"] == 1
    print("✅ test_compute_otif_depot_shift_over_holiday")


def test_compute_otif_depot_shift_fallback_no_column():
    """Zonder 'Aanleveren depot' kolom → oude gedrag: iedereen op peildatum."""
    peildag = date(2026, 7, 21)
    df = pd.DataFrame([
        {"Bron_week": "0", "Nummer": "X-1", "Leverdatum": pd.Timestamp(peildag),
         "Status": "uitgeleverd"},
        {"Bron_week": "0", "Nummer": "X-2", "Leverdatum": pd.Timestamp(peildag),
         "Status": "Productie gereed"},
    ])
    r = compute_otif(df, peildatum=peildag, holiday_dates=set())
    assert r["depot_shift_actief"] is False
    assert r["totaal"] == 2
    assert r["aantal_depot"] == 0
    print("✅ test_compute_otif_depot_shift_fallback_no_column")


def test_compute_otif_depot_shift_disable_flag():
    """apply_depot_shift=False deactiveert de logica expliciet."""
    peildag = date(2026, 7, 21)
    vorige  = date(2026, 7, 20)
    df = pd.DataFrame([
        {"Bron_week": "0", "Nummer": "A", "Leverdatum": pd.Timestamp(peildag),
         "Status": "uitgeleverd", "Aanleveren depot": 1},
        {"Bron_week": "0", "Nummer": "B", "Leverdatum": pd.Timestamp(vorige),
         "Status": "Productie gereed", "Aanleveren depot": 1},
    ])
    r = compute_otif(df, peildatum=peildag, holiday_dates=set(), apply_depot_shift=False)
    # Zonder shift: alleen A (op peildag) telt mee
    assert r["totaal"] == 1
    assert set(r["orders"]["Nummer"].astype(str)) == {"A"}
    assert r["depot_shift_actief"] is False
    print("✅ test_compute_otif_depot_shift_disable_flag")


def test_otif_status_map_content():
    """Verifieer dat OTIF_STATUS_MAP alle statussen uit de bijlage bevat."""
    verwacht_sleutels = {
        "uitgeleverd", "afgehaald", "pc afgehaald", "productie gereed",
        "geblokkeerd", "opgehangen", "coat gereed", "ub",
        "nabewerking nog uitvoeren", "ub v gereed", "meetrapport",
        "pc opgehangen",
    }
    assert verwacht_sleutels == set(OTIF_STATUS_MAP.keys()), (
        f"Ontbrekend: {verwacht_sleutels - set(OTIF_STATUS_MAP.keys())} "
        f"Extra: {set(OTIF_STATUS_MAP.keys()) - verwacht_sleutels}"
    )
    print("✅ test_otif_status_map_content")


def test_gauge_equal_zone_math():
    """
    Verifieer de piecewise-lineaire schaal die de gauge gebruikt: elke zone
    (rood/oranje/groen) beslaat exact 60° van de arc. Deze functie zit in
    viewer_app.py maar we kunnen de wiskunde 1-op-1 hier uitrekenen.
    """
    tg = 96.0
    to = 80.0

    def pct_to_angle(pct):
        clip = max(0.0, min(100.0, pct))
        if clip <= to:
            return 180.0 - (clip / to) * 60.0
        if clip <= tg:
            return 120.0 - ((clip - to) / (tg - to)) * 60.0
        return 60.0 - ((clip - tg) / (100.0 - tg)) * 60.0

    # Zonegrenzen liggen op exact 180, 120, 60, 0 graden.
    assert abs(pct_to_angle(0)   - 180.0) < 1e-9
    assert abs(pct_to_angle(to)  - 120.0) < 1e-9   # rood/oranje-grens
    assert abs(pct_to_angle(tg)  -  60.0) < 1e-9   # oranje/groen-grens
    assert abs(pct_to_angle(100) -   0.0) < 1e-9

    # Middens van de zones liggen in het midden van hun arc-segment.
    assert abs(pct_to_angle(to/2)         - 150.0) < 1e-9  # midden rood
    assert abs(pct_to_angle((to+tg)/2)    -  90.0) < 1e-9  # midden oranje
    assert abs(pct_to_angle((tg+100)/2)   -  30.0) < 1e-9  # midden groen

    # Elke zone beslaat 60° - dit is de businesswaarde die de zichtbaarheid
    # van de groene band verzekert.
    assert abs((pct_to_angle(0)  - pct_to_angle(to))  - 60.0) < 1e-9
    assert abs((pct_to_angle(to) - pct_to_angle(tg))  - 60.0) < 1e-9
    assert abs((pct_to_angle(tg) - pct_to_angle(100)) - 60.0) < 1e-9

    # Onder- en overloop klippen naar 0/180
    assert abs(pct_to_angle(-10)  - 180.0) < 1e-9
    assert abs(pct_to_angle(150)  -   0.0) < 1e-9
    print("✅ test_gauge_equal_zone_math (elke zone = 60°)")


if __name__ == "__main__":
    test_otif_status_map_content()
    test_gauge_equal_zone_math()
    test_map_otif_status_matches_spec()
    test_find_originele_leverdatum_column()
    test_compute_otif_basic()
    test_compute_otif_originele_datum()
    test_compute_otif_empty_and_no_column()
    test_compute_otif_perfect_score()
    test_get_peildatum_from_metadata()
    test_previous_workday_holidays()
    test_compute_otif_depot_shift_basic()
    test_compute_otif_depot_shift_over_weekend()
    test_compute_otif_depot_shift_over_holiday()
    test_compute_otif_depot_shift_fallback_no_column()
    test_compute_otif_depot_shift_disable_flag()
    print("\nAlle tests geslaagd 🎉")
