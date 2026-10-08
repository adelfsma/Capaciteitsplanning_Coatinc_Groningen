"""Tests voor bereken_eerstvolgende_leverdatum (v2.5.15)."""
from datetime import date

import pandas as pd

from shared import bereken_eerstvolgende_leverdatum

CAP = 60_000  # kg per dag
VANDAAG = date(2026, 10, 8)  # donderdag


def _df(belasting: dict) -> pd.DataFrame:
    """belasting: {'YYYY-MM-DD': benutting_pct} -> open orders-DataFrame."""
    rows = [
        {
            "Verzinkdatum": pd.Timestamp(d),
            "Verzinkstatus": "Niet verzinkt",
            "Reden_uitsluiting": "",
            "Gewicht_effectief_kg": pct / 100 * CAP,
        }
        for d, pct in belasting.items()
    ]
    return pd.DataFrame(rows)


def _hol(*datums):
    return pd.DataFrame({"Datum": [pd.Timestamp(d).date() for d in datums]})


def _run(df, hol=None, drempel=95):
    return bereken_eerstvolgende_leverdatum(
        df, hol if hol is not None else _hol(), CAP, drempel, VANDAAG
    )


def test_leeg_standaard_levertijd():
    r = _run(_df({}))
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-13")  # werkdag 3
    assert r["leverdatum"] == pd.Timestamp("2026-10-15")    # werkdag 5
    assert r["aantal_opgeschoven"] == 0


def test_voorbeeld_uit_voorstel():
    df = _df({
        "2026-10-13": 92, "2026-10-14": 98, "2026-10-15": 94,
        "2026-10-16": 90, "2026-10-19": 85,
    })
    r = _run(df)
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-15")
    assert r["leverdatum"] == pd.Timestamp("2026-10-19")
    assert round(r["buffer_gemiddelde"], 1) == 87.5


def test_d_zelf_telt_niet_mee_in_gemiddelde():
    # D=94, D+1=96, D+2=93 -> gemiddelde 94.5 < 95 -> akkoord
    r = _run(_df({"2026-10-13": 94, "2026-10-14": 96, "2026-10-15": 93}))
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-13")


def test_grens_95_is_niet_akkoord():
    r = _run(_df({"2026-10-13": 95}))
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-14")


def test_feestdag_overgeslagen():
    # di 13-10 feestdag: vroegste verzinkdag schuift naar wo 14-10,
    # buffer- en leverdagen slaan de feestdag ook over.
    r = _run(_df({}), _hol("2026-10-13"))
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-14")
    assert r["leverdatum"] == pd.Timestamp("2026-10-16")


def test_coat_en_verzinkt_tellen_niet_mee():
    df = _df({"2026-10-13": 100})
    df = pd.concat([df.assign(Reden_uitsluiting="Coat-order"),
                    df.assign(Verzinkstatus="Verzinkt")])
    r = _run(df)
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-13")


def test_niets_gevonden():
    vol = {str(d.date()): 100 for d in pd.bdate_range("2026-10-09", periods=40)}
    r = _run(_df(vol))
    assert r["gevonden"] is False and r["leverdatum"] is None


def test_drempel_slider():
    df = _df({"2026-10-13": 88})
    assert _run(df, drempel=85)["verzinkdatum"] == pd.Timestamp("2026-10-14")
    assert _run(df, drempel=90)["verzinkdatum"] == pd.Timestamp("2026-10-13")
