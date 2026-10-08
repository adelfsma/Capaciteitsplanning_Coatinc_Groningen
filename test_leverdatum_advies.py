"""Tests voor eerstvolgende leverdatum en balie-capaciteit (v2.5.22)."""
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
    assert r["levertijd_werkdagen"] == 5


def test_piek_wordt_teruggeschoven():
    # vr 16-10 140%: 45% boven de drempel schuift terug naar 15, 14 en 13-10
    # (elk 80% -> 95%). Daarmee zijn 13 t/m 16-10 vol; eerste vrije dag ma 19-10.
    df = _df({"2026-10-13": 80, "2026-10-14": 80, "2026-10-15": 80, "2026-10-16": 140})
    r = _run(df)
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-19")
    assert r["leverdatum"] == pd.Timestamp("2026-10-21")
    assert r["aantal_opgeschoven"] == 4  # 13, 14, 15 en 16-10 vol


def test_terugschuiven_max_3_werkdagen():
    # vr 16-10 150%: 55% over. 15/14/13-10 (90%) nemen elk 5% op, de
    # resterende 40% blijft op 16-10. Ma 12-10 (4 werkdagen terug) krijgt niets.
    df = _df({"2026-10-12": 80, "2026-10-13": 90, "2026-10-14": 90, "2026-10-15": 90, "2026-10-16": 150})
    r = _run(df)
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-19")
    rij = _tabel(df, _hol(), "2026-10-20", altijd_vol_werkdagen=0)
    assert rij.loc["14-10", "Verzinkdatum"] == pd.Timestamp("2026-10-12")
    assert rij.loc["14-10", "Klasse_key"] == "k4"  # 12-10 blijft 80% -> 12.000 kg vrij


def test_piek_schuift_niet_naar_later():
    # Piek op 19-10 mag niet naar 20-10 (dat zou te laat zijn): 20-10 blijft vrij
    # 13-10 95% (vol), 14-16-10 94% -> na terugschuiven 95% (vol), 19-10 houdt 137%
    df = _df({"2026-10-13": 95, "2026-10-14": 94, "2026-10-15": 94, "2026-10-16": 94, "2026-10-19": 140})
    r = _run(df)
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-20")


def test_piek_schuift_niet_naar_verleden():
    # Piek op vr 09-10 kan alleen naar do 08-10 (vandaag), niet naar 07-10
    df = _df({"2026-10-07": 0, "2026-10-08": 0, "2026-10-09": 250})
    rij = _tabel(df, _hol(), "2026-10-13", altijd_vol_werkdagen=0)
    # leverdatum 12-10 <- verzinkdag 08-10: vol door teruggeschoven piek
    assert rij.loc["12-10", "Klasse"] == "Vol"


def test_grens_95_is_niet_akkoord():
    r = _run(_df({"2026-10-13": 95}))
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-14")


def test_feestdag_overgeslagen():
    # di 13-10 feestdag: vroegste verzinkdag schuift naar wo 14-10,
    # buffer- en leverdagen slaan de feestdag ook over.
    r = _run(_df({}), _hol("2026-10-13"))
    assert r["verzinkdatum"] == pd.Timestamp("2026-10-14")
    assert r["leverdatum"] == pd.Timestamp("2026-10-16")
    assert r["levertijd_werkdagen"] == 5  # feestdag telt niet als werkdag


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
    assert capaciteit_klasse(0)[1] == "vol"
    assert capaciteit_klasse(999)[1] == "vol"
    assert capaciteit_klasse(-500)[1] == "vol"
    assert capaciteit_klasse(1_000)[1] == "k2"
    assert capaciteit_klasse(4_999)[1] == "k2"
    assert capaciteit_klasse(5_000)[1] == "k3"
    assert capaciteit_klasse(10_000)[1] == "k4"


def _tabel(df, hol, tot, **kw):
    t = bereken_balie_capaciteit(df, hol, CAP, VANDAAG, pd.Timestamp(tot), **kw)
    return t.set_index(t["Datum"].dt.strftime("%d-%m"))


# De balietabel staat per LEVERDATUM; capaciteit = die van de verzinkdag
# (leverdatum − 2 werkdagen). _df() zet belasting op de verzinkdag.

def test_eerste_vijf_werkdagen_niet_beschikbaar():
    # Lege planning, vandaag do 08-10: leverdatums 08 t/m 14-10 niet beschikbaar,
    # do 15-10 (= vandaag + 5 werkdagen) is de eerste leverbare dag.
    rij = _tabel(_df({}), _hol(), "2026-10-16")
    for dag in ["08-10", "09-10", "12-10", "13-10", "14-10"]:
        assert rij.loc[dag, "Klasse"] == "Niet beschikbaar"
    assert rij.loc["15-10", "Klasse_key"] == "k4"
    # leverdatum = verzinkdatum + 2 werkdagen
    assert rij.loc["15-10", "Verzinkdatum"] == pd.Timestamp("2026-10-13")
    # Instelbaar: 0 = uit
    rij0 = _tabel(_df({}), _hol(), "2026-10-16", altijd_vol_werkdagen=0)
    assert rij0.loc["08-10", "Klasse_key"] == "k4"


def test_feestdag_bij_niet_beschikbaar():
    # vr 09-10 feestdag -> 5 werkdagen na do 08-10 is vr 16-10
    rij = _tabel(_df({}), _hol("2026-10-09"), "2026-10-19")
    assert rij.loc["09-10", "Klasse"] == "Gesloten"
    assert rij.loc["15-10", "Klasse"] == "Niet beschikbaar"
    assert rij.loc["16-10", "Klasse_key"] == "k4"
    # leverdatum di 13-10 -> verzinkdag 2 werkdagen terug, feestdag overgeslagen: do 08-10
    assert rij.loc["13-10", "Verzinkdatum"] == pd.Timestamp("2026-10-08")


def test_weekend_sluit_aan_op_advies():
    # Vandaag zaterdag 10-10: vandaag + 5 werkdagen = vr 16-10 = leverdatum uit het advies
    t = bereken_balie_capaciteit(_df({}), _hol(), CAP, date(2026, 10, 10), pd.Timestamp("2026-10-16"))
    rij = t.set_index(t["Datum"].dt.strftime("%d-%m"))
    assert rij.loc["15-10", "Klasse"] == "Niet beschikbaar"
    assert rij.loc["16-10", "Klasse_key"] == "k4"
    adv = bereken_eerstvolgende_leverdatum(_df({}), _hol(), CAP, 95, date(2026, 10, 10))
    assert adv["leverdatum"] == pd.Timestamp("2026-10-16")


def test_balie_tabel_per_leverdatum():
    # verzinkdag wo 14-10 overboekt (102%) -> leverdatum vr 16-10 Vol
    # verzinkdag do 15-10 94% (3.600 kg vrij) -> leverdatum ma 19-10 1.000-5.000
    df = _df({"2026-10-14": 102, "2026-10-15": 94})
    rij = _tabel(df, _hol("2026-10-13"), "2026-10-20")
    assert list(rij.index) == ["08-10", "09-10", "12-10", "13-10", "14-10", "15-10", "16-10", "19-10", "20-10"]
    assert rij.loc["13-10", "Klasse"] == "Gesloten" and not bool(rij.loc["13-10", "Vol"])
    assert rij.loc["16-10", "Klasse"] == "Vol" and bool(rij.loc["16-10", "Vol"])
    assert rij.loc["19-10", "Klasse_key"] == "k2"
    assert rij.loc["20-10", "Klasse_key"] == "k4"


def test_vol_volgt_drempel():
    # Precies 95% is Vol (zelfde grens als het advies: alleen < 95% is beschikbaar)
    # verzinkdag di 13-10 95% -> leverdatum do 15-10; verzinkdag wo 14-10 90% -> vr 16-10
    df = _df({"2026-10-13": 95, "2026-10-14": 90})
    rij = _tabel(df, _hol(), "2026-10-16")
    assert rij.loc["15-10", "Klasse"] == "Vol"
    assert rij.loc["16-10", "Klasse_key"] == "k3"  # 10% van 60 ton = 6.000 kg vrij
    rij90 = _tabel(df, _hol(), "2026-10-16", vol_drempel_pct=90)
    assert rij90.loc["16-10", "Klasse"] == "Vol"


def test_export_0810_60_ton():
    # Situatie uit de export van 08-10 bij 60 ton (benutting t.o.v. capaciteit).
    # Piek ma 19-10 (137%) schuift terug: 16-10 -> 95%, 15-10 -> 95%, rest naar 14-10.
    benut = {
        "2026-10-12": 78.7, "2026-10-13": 81.1, "2026-10-14": 69.3, "2026-10-15": 82.7,
        "2026-10-16": 70.2, "2026-10-19": 137.2, "2026-10-20": 70.0, "2026-10-21": 88.0,
        "2026-10-22": 51.2,
    }
    df = _df(benut)
    rij = _tabel(df, _hol(), "2026-10-26")
    assert rij.loc["15-10", "Klasse_key"] == "k4"  # verzinkdag 13-10: 81%
    assert rij.loc["16-10", "Klasse_key"] == "k4"  # verzinkdag 14-10: 69% + rest piek = 74%
    assert rij.loc["19-10", "Klasse"] == "Vol"     # verzinkdag 15-10: 95% na terugschuiven
    assert rij.loc["20-10", "Klasse"] == "Vol"     # verzinkdag 16-10: 95%
    assert rij.loc["21-10", "Klasse"] == "Vol"     # verzinkdag 19-10: 95%
    assert rij.loc["22-10", "Klasse_key"] == "k4"  # verzinkdag 20-10: 70%
    adv = _run(df)
    assert adv["leverdatum"] == pd.Timestamp("2026-10-15")
    beschikbaar = rij[~rij["Vol"] & ~rij["Gesloten"]]
    assert beschikbaar.index[0] == "15-10"
