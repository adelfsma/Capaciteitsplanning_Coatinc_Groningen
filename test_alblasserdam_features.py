"""
Tests voor de Alblasserdam-specifieke features:
1. Poetsen-uitzondering op OTIF: voor orders waarvan PrijsCategorie 'poetsen'
   bevat, telt de status 'Afgehaald' als niet OK ('Nee') i.p.v. OK.
2. Scheepsleidingen-klanten (24-uursservice) worden als 'Aanleveren depot = 1'
   behandeld (gematcht op klantnummer).

Beide features zijn standaard uit / zonder effect, zodat Groningen ongewijzigd
blijft. De tests raken alleen pure helpers uit shared.py.
"""
import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))

from shared import (  # noqa: E402
    compute_otif,
    _load_scheepsleidingen_ids,
    _force_scheepsleidingen_depot,
    get_poetsen_otif_actief,
    SCHEEPSLEIDINGEN_FILE,
)


def _poetsen_df(peildag: date) -> pd.DataFrame:
    """Vier orders op de peildag met een mix van poetsen/niet-poetsen en
    statussen 'Afgehaald'/'Coat gereed'."""
    ts = pd.Timestamp(peildag)
    return pd.DataFrame([
        # Poetsen + Afgehaald → moet 'Nee' worden met de vlag aan
        {"Bron_week": "0", "Nummer": "P-001", "Leverdatum": ts, "Status": "Afgehaald",
         "PrijsCategorie": "Verzinken 2D,Poetsen licht ( B )"},
        # Poetsen + Coat gereed → blijft 'Ja'
        {"Bron_week": "0", "Nummer": "P-002", "Leverdatum": ts, "Status": "Coat gereed",
         "PrijsCategorie": "Verzinken 1D,Poetsen zwaar"},
        # NIET-poetsen + Afgehaald → blijft 'Ja' (normale regel)
        {"Bron_week": "0", "Nummer": "P-003", "Leverdatum": ts, "Status": "Afgehaald",
         "PrijsCategorie": "Verzinken 2D licht"},
        # Poetsen + Productie gereed → 'Nee' (was al 'Nee', ongewijzigd)
        {"Bron_week": "0", "Nummer": "P-004", "Leverdatum": ts, "Status": "Productie gereed",
         "PrijsCategorie": "Poetsen ( B )"},
    ])


def test_poetsen_uit_geen_effect():
    """Vlag uit → normale OTIF-regels: Afgehaald telt als OK, ook voor poetsen."""
    peildag = date(2026, 7, 20)
    df = _poetsen_df(peildag)
    r = compute_otif(df, peildatum=peildag, poetsen_afgehaald_niet_ok=False)
    # Ja: P-001 (Afgehaald), P-002 (Coat gereed), P-003 (Afgehaald) = 3
    assert r["gereed"] == 3, f"gereed = {r['gereed']}"
    # Nee: P-004 (Productie gereed) = 1
    assert r["niet_gereed"] == 1, f"niet_gereed = {r['niet_gereed']}"
    print("✅ test_poetsen_uit_geen_effect")


def test_poetsen_aan_afgehaald_niet_ok():
    """Vlag aan → poetsen + Afgehaald wordt 'Nee'; niet-poetsen + Afgehaald blijft 'Ja'."""
    peildag = date(2026, 7, 20)
    df = _poetsen_df(peildag)
    r = compute_otif(df, peildatum=peildag, poetsen_afgehaald_niet_ok=True)
    # Ja: P-002 (poetsen Coat gereed), P-003 (niet-poetsen Afgehaald) = 2
    assert r["gereed"] == 2, f"gereed = {r['gereed']}"
    # Nee: P-001 (poetsen Afgehaald → geflipt), P-004 (Productie gereed) = 2
    assert r["niet_gereed"] == 2, f"niet_gereed = {r['niet_gereed']}"

    orders = r["orders"].set_index("Nummer")["OTIF_status"].to_dict()
    assert orders["P-001"] == "Nee", "poetsen Afgehaald moet Nee zijn"
    assert orders["P-002"] == "Ja", "poetsen Coat gereed moet Ja blijven"
    assert orders["P-003"] == "Ja", "niet-poetsen Afgehaald moet Ja blijven"
    assert orders["P-004"] == "Nee", "poetsen Productie gereed blijft Nee"
    # Formule: (1 - 2/4) * 100 = 50%
    assert abs(r["otif_pct"] - 50.0) < 1e-6, f"otif_pct = {r['otif_pct']}"
    print("✅ test_poetsen_aan_afgehaald_niet_ok")


def test_poetsen_aan_zonder_prijscategorie_kolom():
    """Vlag aan maar geen PrijsCategorie-kolom → mag niet crashen; normale regels."""
    peildag = date(2026, 7, 20)
    df = _poetsen_df(peildag).drop(columns=["PrijsCategorie"])
    r = compute_otif(df, peildatum=peildag, poetsen_afgehaald_niet_ok=True)
    # Zonder PrijsCategorie kan de uitzondering niet toegepast worden → Afgehaald = Ja
    assert r["gereed"] == 3, f"gereed = {r['gereed']}"
    assert r["niet_gereed"] == 1, f"niet_gereed = {r['niet_gereed']}"
    print("✅ test_poetsen_aan_zonder_prijscategorie_kolom")


def test_poetsen_hoofdletterongevoelig():
    """'POETSEN' / 'poetsen' / gemengd moeten allemaal matchen."""
    peildag = date(2026, 7, 20)
    df = pd.DataFrame([
        {"Bron_week": "0", "Nummer": "X1", "Leverdatum": pd.Timestamp(peildag),
         "Status": "Afgehaald", "PrijsCategorie": "iets POETSEN iets"},
        {"Bron_week": "0", "Nummer": "X2", "Leverdatum": pd.Timestamp(peildag),
         "Status": "Afgehaald", "PrijsCategorie": "PoEtSeN"},
    ])
    r = compute_otif(df, peildatum=peildag, poetsen_afgehaald_niet_ok=True)
    assert r["niet_gereed"] == 2, f"niet_gereed = {r['niet_gereed']}"
    print("✅ test_poetsen_hoofdletterongevoelig")


def test_scheepsleidingen_ids_missing_file(tmp_path):
    """Ontbrekend bestand → lege set (geen effect)."""
    assert _load_scheepsleidingen_ids(tmp_path) == set()
    print("✅ test_scheepsleidingen_ids_missing_file")


def test_scheepsleidingen_ids_alleen_actief(tmp_path):
    """Alleen actieve klanten tellen mee; nummers worden genormaliseerd."""
    df = pd.DataFrame({
        "is_active": ["Active", "active", "Inactive"],
        "number": [12161705, "12161725", 99999999],
        "name": ["A", "B", "C"],
    })
    df.to_excel(tmp_path / SCHEEPSLEIDINGEN_FILE, index=False)
    ids = _load_scheepsleidingen_ids(tmp_path)
    assert ids == {"12161705", "12161725"}, f"ids = {ids}"
    assert "99999999" not in ids, "inactieve klant mag niet meetellen"
    print("✅ test_scheepsleidingen_ids_alleen_actief")


def test_force_scheepsleidingen_depot():
    """Order-regels van scheepsleidingen-klanten krijgen Aanleveren depot = 1."""
    order = pd.DataFrame({
        "Ordernummer": [1, 2, 3],
        "ID Debiteur": ["12161705", "12160497", "12161725"],
        "Aanleveren depot": [0, 0, 0],
    })
    out = _force_scheepsleidingen_depot(order.copy(), {"12161705", "12161725"})
    depot = dict(zip(out["ID Debiteur"], out["Aanleveren depot"]))
    assert depot["12161705"] == 1, "scheepsleidingen-klant moet depot=1 krijgen"
    assert depot["12161725"] == 1, "scheepsleidingen-klant moet depot=1 krijgen"
    assert depot["12160497"] == 0, "gewone klant blijft depot=0"
    print("✅ test_force_scheepsleidingen_depot")


def test_force_scheepsleidingen_depot_lege_set():
    """Lege set of ontbrekende kolom → ongewijzigd."""
    order = pd.DataFrame({"ID Debiteur": ["1"], "Aanleveren depot": [0]})
    out = _force_scheepsleidingen_depot(order.copy(), set())
    assert out["Aanleveren depot"].tolist() == [0]
    # Ontbrekende depot-kolom → geen crash
    order2 = pd.DataFrame({"ID Debiteur": ["1"]})
    out2 = _force_scheepsleidingen_depot(order2.copy(), {"1"})
    assert "Aanleveren depot" not in out2.columns
    print("✅ test_force_scheepsleidingen_depot_lege_set")


def test_poetsen_vlag_default_uit():
    """Zonder [locatie]-secret is de poetsen-uitzondering uit (Groningen)."""
    assert get_poetsen_otif_actief() is False
    print("✅ test_poetsen_vlag_default_uit")


if __name__ == "__main__":
    # Tests met een tmp_path-fixture draaien via pytest; hier alleen de tests
    # zonder fixture voor een snelle rooktest.
    test_poetsen_uit_geen_effect()
    test_poetsen_aan_afgehaald_niet_ok()
    test_poetsen_aan_zonder_prijscategorie_kolom()
    test_poetsen_hoofdletterongevoelig()
    test_force_scheepsleidingen_depot()
    test_force_scheepsleidingen_depot_lege_set()
    test_poetsen_vlag_default_uit()
    print("Rooktest OK. Volledige suite: python -m pytest test_alblasserdam_features.py -q")
