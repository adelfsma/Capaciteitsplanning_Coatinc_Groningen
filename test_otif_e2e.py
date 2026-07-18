"""
End-to-end scenariotest:
- Bouwt een 'merged'-DataFrame zoals load_published_data zou produceren
  (Bron_week + Status + Leverdatum + Originele_leverdatum).
- Draait de volledige OTIF-computatie zoals viewer_app.py doet.
- Controleert de Excel-serialisatie van de OTIF-tabel.
"""
from datetime import date
import io

import pandas as pd

from shared import (
    OTIF_GAUGE_GREEN_THRESHOLD,
    OTIF_GAUGE_ORANGE_THRESHOLD,
    compute_otif,
    get_peildatum_from_metadata,
    map_otif_status,
    find_originele_leverdatum_column,
)


def build_realistic_merged(peildag: date, verschuiving_dagen: int = 3) -> pd.DataFrame:
    """
    Simuleer een merged-DataFrame met de kolommen die de viewer gebruikt.
    Verschuivingsdagen: hoeveel dagen de 'Originele_leverdatum' voor sommige
    orders verschoven is t.o.v. de 'Leverdatum' (om beide varianten te toetsen).
    """
    peil_ts = pd.Timestamp(peildag)
    verschoven_ts = peil_ts - pd.Timedelta(days=verschuiving_dagen)

    rijen = []

    # 60 orders op de peildag met een realistische mix
    statussen = (
        ["uitgeleverd"] * 25
        + ["Afgehaald"] * 10
        + ["coat gereed"] * 5
        + ["UB V Gereed"] * 3
        + ["Productie gereed"] * 6
        + ["Geblokkeerd"] * 2
        + ["Opgehangen"] * 4
        + ["PC Afgehaald"] * 3
        + ["Gereserveerd*"] * 2
    )
    for i, status in enumerate(statussen):
        # De 'Originele_leverdatum' is voor de helft van de nee-orders verschoven.
        oud = peil_ts
        if status == "Productie gereed" and i % 2 == 0:
            oud = verschoven_ts
        rijen.append({
            "Bronbestand": "Export.xlsx",
            "Bron_week": "0",
            "Nummer": f"ORD-{1000+i}",
            "Ordernummer_base": 1000 + i,
            "Debiteurnummer": 90000 + (i % 10),
            "Klantnaam": f"Klant {(i % 12) + 1}",
            "Segment_debtor_export": "",
            "Materiaaltype": "Overig / onbekend",
            "Datum": peil_ts,
            "Leverdatum": peil_ts,
            "Originele_leverdatum": oud,
            "Status": status,
            "Verzinkstatus": "Verzinkt",
            "Gewicht_effectief_kg": 300.0 + i * 5,
            "Gewicht_bron": "Export+",
        })

    # 15 reserveringen op de peildag (moeten uitgesloten worden!)
    for i in range(15):
        rijen.append({
            "Bronbestand": "OrderExport2G.xlsx",
            "Bron_week": "reservering",
            "Nummer": f"RES-{2000+i}",
            "Ordernummer_base": None,
            "Debiteurnummer": 91000 + i,
            "Klantnaam": f"Reserveringklant {i}",
            "Segment_debtor_export": "",
            "Materiaaltype": "Overig / onbekend",
            "Datum": peil_ts,
            "Leverdatum": peil_ts,
            "Originele_leverdatum": pd.NaT,
            "Status": "Reservering",
            "Verzinkstatus": "Niet verzinkt",
            "Gewicht_effectief_kg": 500.0,
            "Gewicht_bron": "Reservering",
        })

    # 20 orders op andere datums (moeten uitgesloten worden)
    for i, offset in enumerate([-2, -1, 1, 2, 3] * 4):
        rijen.append({
            "Bronbestand": "Export+1.xlsx",
            "Bron_week": "+1",
            "Nummer": f"OTH-{3000+i}",
            "Ordernummer_base": 3000 + i,
            "Debiteurnummer": 92000 + i,
            "Klantnaam": f"Andere klant {i}",
            "Segment_debtor_export": "",
            "Materiaaltype": "Overig / onbekend",
            "Datum": peil_ts + pd.Timedelta(days=offset),
            "Leverdatum": peil_ts + pd.Timedelta(days=offset),
            "Originele_leverdatum": peil_ts + pd.Timedelta(days=offset),
            "Status": "Productie gereed",
            "Verzinkstatus": "Niet verzinkt",
            "Gewicht_effectief_kg": 400.0,
            "Gewicht_bron": "Export+",
        })

    return pd.DataFrame(rijen)


def test_realistic_scenario():
    peildag = date(2026, 7, 17)
    merged = build_realistic_merged(peildag, verschuiving_dagen=3)

    print(f"Test-scenario: {len(merged)} rijen, peildatum {peildag}\n")

    # === Variant 1: OTIF t.o.v. Leverdatum (= Datum) ===
    r1 = compute_otif(merged, peildatum=peildag, date_column="Leverdatum", date_column_label="Datum")
    print("Variant 1 – Datum:")
    print(f"  totaal={r1['totaal']}, ja={r1['gereed']}, nee={r1['niet_gereed']}, "
          f"nvt={r1['nvt']}, onb={r1['onbekend']}, OTIF={r1['otif_pct']:.2f}%")

    # Handmatige controle:
    #   Ja: 25+10+5+3=43
    #   Nee: 6+2+4=12
    #   Nvt: 3
    #   Onbekend: 2 (Gereserveerd*)
    #   Totaal (excl. reserveringen): 60
    #   OTIF = (1 - 12/60) * 100 = 80%
    assert r1["totaal"] == 60
    assert r1["gereed"] == 43
    assert r1["niet_gereed"] == 12
    assert r1["nvt"] == 3
    assert r1["onbekend"] == 2
    assert abs(r1["otif_pct"] - 80.0) < 1e-6

    # === Variant 2: OTIF t.o.v. Originele_leverdatum ===
    # Enkele 'Productie gereed' orders hebben Originele_leverdatum verschoven,
    # dus die vallen bij deze variant NIET meer in de subset van peildag.
    r2 = compute_otif(merged, peildatum=peildag, date_column="Originele_leverdatum",
                      date_column_label="Originele datum")
    print("\nVariant 2 – Originele datum:")
    print(f"  totaal={r2['totaal']}, ja={r2['gereed']}, nee={r2['niet_gereed']}, "
          f"nvt={r2['nvt']}, onb={r2['onbekend']}, OTIF={r2['otif_pct']:.2f}%")

    # 3 'Productie gereed' orders zijn verschoven (i=[0,2,4] van de 6, dus 3 stuks)
    # Dus 60 - 3 = 57 in de subset, met 12 - 3 = 9 nee.
    # OTIF = (1 - 9/57) * 100 ≈ 84.21%
    assert r2["totaal"] == 57
    assert r2["gereed"] == 43
    assert r2["niet_gereed"] == 9
    assert abs(r2["otif_pct"] - (1 - 9/57) * 100) < 1e-6

    # === Peildatum-mechanisme ===
    peildatum_uit_meta = get_peildatum_from_metadata({"published_at": "17-07-2026 09:15:22"})
    assert peildatum_uit_meta == peildag

    # === Excel-serialisatie ===
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        r1["orders"].to_excel(writer, index=False, sheet_name="OTIF-orders")
        pd.DataFrame([{"OTIF (%)": round(r1["otif_pct"], 2),
                       "Peildatum": str(peildag),
                       "Vergelijking": r1["date_column_label"]}]).to_excel(
            writer, index=False, sheet_name="Samenvatting"
        )
    buf.seek(0)
    print(f"\nExcel-bestand grootte: {len(buf.getvalue())} bytes ✅")

    # === Kolomdetectie ===
    kolom = find_originele_leverdatum_column(merged)
    print(f"Detectie Originele_leverdatum-kolom in merged: {kolom}")

    print(f"\nGroene drempel: {OTIF_GAUGE_GREEN_THRESHOLD}%  |  "
          f"Oranje drempel: {OTIF_GAUGE_ORANGE_THRESHOLD}%")
    print("\n✅ Alle e2e-scenariocontroles geslaagd 🎉")


if __name__ == "__main__":
    test_realistic_scenario()
