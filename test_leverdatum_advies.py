"""Tests voor eerstvolgende leverdatum en balie-capaciteit (v2.5.16)."""
from datetime import date

import pandas as pd

from shared import (
    bereken_balie_capaciteit,
    bereken_eerstvolgende_leverdatum,
    capaciteit_klasse,
)

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


def test_voorbeeld_buffer_3_dagen():
    # D=13: buffer 14/15/16 = 98/97/94 -> gem 96,3 -> schuift
    # D=14 (98) en D=15 (97) vallen zelf al af
    # D=16: 94, buffer 19/20/21 = 85/0/0 -> akkoord
    df = _df({
        "2026-10-13": 92, "2026-10-14": 98, "2026-10-15": 97,
        "2026-10-16": 94, "2026-10-19": 85,
    })
    r = _run(df)
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-16")
    assert r["leverdatum"] == pd.Timestamp("2026-10-20")
    assert r["buffer_dagen"] == (
        pd.Timestamp("2026-10-19"), pd.Timestamp("2026-10-20"), pd.Timestamp("2026-10-21"),
    )
    assert r["aantal_opgeschoven"] == 3


def test_d_zelf_telt_niet_mee_in_gemiddelde():
    # D=94, buffer 96/96/92 -> gemiddelde 94,7 < 95 -> akkoord
    r = _run(_df({"2026-10-13": 94, "2026-10-14": 96, "2026-10-15": 96, "2026-10-16": 92}))
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-13")


def test_derde_bufferdag_telt_mee():
    # Met 2 bufferdagen zou D=13 akkoord zijn (gem 90); de derde dag (99)
    # trekt het gemiddelde naar 93 -> nog akkoord; bij 105 -> 95 -> schuift.
    assert _run(_df({"2026-10-14": 90, "2026-10-15": 90, "2026-10-16": 99}))["verzinkdatum"] \
        == pd.Timestamp("2026-10-13")
    assert _run(_df({"2026-10-14": 90, "2026-10-15": 90, "2026-10-16": 105}))["verzinkdatum"] \
        != pd.Timestamp("2026-10-13")


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


# ── Balie ────────────────────────────────────────────────────────────────────

def test_capaciteit_klassen():
    assert capaciteit_klasse(0)[1] == "k1"
    assert capaciteit_klasse(999)[1] == "k1"
    assert capaciteit_klasse(1_000)[1] == "k2"
    assert capaciteit_klasse(4_999)[1] == "k2"
    assert capaciteit_klasse(5_000)[1] == "k3"
    assert capaciteit_klasse(10_000)[1] == "k4"
    assert capaciteit_klasse(-500)[1] == "k1"


def test_balie_tabel():
    # di 13-10 feestdag; vr 09-10 overboekt (102%); ma 12-10 heeft 2.400 kg vrij
    df = _df({"2026-10-09": 102, "2026-10-12": 96})
    t = bereken_balie_capaciteit(df, _hol("2026-10-13"), CAP, VANDAAG, pd.Timestamp("2026-10-15"))
    assert list(t["Datum"].dt.strftime("%d-%m")) == ["08-10", "09-10", "12-10", "13-10", "14-10", "15-10"]
    rij = t.set_index(t["Datum"].dt.strftime("%d-%m"))
    assert rij.loc["09-10", "Beschikbaar_kg"] == 0 and rij.loc["09-10", "Klasse_key"] == "k1"
    assert rij.loc["12-10", "Klasse_key"] == "k2"
    assert bool(rij.loc["13-10", "Gesloten"]) and rij.loc["13-10", "Klasse"] == "Gesloten"
    assert rij.loc["14-10", "Klasse_key"] == "k4"
