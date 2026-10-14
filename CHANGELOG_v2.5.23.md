# v2.5.23 — Stuksplanning (aantal te verzinken stuks per dag)

## Overzicht

Naast de planning in kg kan de app nu ook plannen op het **aantal te verzinken
stuks per verzinkdag**. Lichter materiaal betekent meer stuks per ton en dus
meer afhaalwerk bij Wit. Een dag kan op gewicht nog ruimte hebben terwijl hij
op stuks al vol zit; dat wordt nu zichtbaar.

De functie staat standaard **uit**. Groningen blijft ongewijzigd.

## Aanzetten (Streamlit-secrets, viewer én beheer)

```toml
[locatie]
stuks_planning = true
# optioneel:
stuks_capaciteit_min = 1000
stuks_capaciteit_max = 8000
stuks_capaciteit_default = 3200   # ≈ 52 traverses × ≈ 61 stuks (CAL 2026, O+M)
stuks_capaciteit_step = 100
kg_per_stuk_default = 34          # schatting bij klanten zonder historie
```

## Nieuw bronbestand: `Orderregels.xlsx`

- De 2G orderregel-export. Hij verschijnt in de beheeromgeving als optioneel
  uploadvak zodra `stuks_planning = true`.
- Benodigde kolommen: `OrderID`, `Subbonnummer V`, `aantal`. Bij het publiceren
  wordt dit gecontroleerd.
- Stuks per deelorder = de som van `aantal` per `Subbonnummer V`. Dit sluit aan
  op `Nummer` in de weekexports (99,9% match op de CAL-export van 14-10-2026).
  Regels zonder `Subbonnummer V` (alleen coat) tellen niet mee.

## Rekenregels

1. **Bevestigde orders**: het werkelijke aantal uit de orderregels.
2. **Reserveringen en regels zonder orderregels**: een schatting, namelijk
   gewicht ÷ kg-per-stuk.
   - Kg-per-stuk van de klant = de mediaan over zijn **gewogen** deelorders in
     de weekexports (`Gewicht` > 0 en een aantal uit de orderregels). Daarvoor
     zijn minimaal 3 deelorders nodig, en de waarde wordt begrensd op 1–1.000
     kg/stuk.
   - Het 2G-ordergewicht wordt hiervoor bewust niet gebruikt. Dat is vaak een
     placeholder (10 kg, 1 ton) en gaf ratio's van 0,005 kg/stuk.
   - Een klant met te weinig historie krijgt `kg_per_stuk_default` (34).
3. Dezelfde uitsluitingen als de kg-planning: verzinkt, UB, coat-orders en
   regels vóór de startdatum tellen niet mee.

## In de viewer

- Zijbalk: slider **Max stuks per dag**.
- Dashboard: grafiek **Stuks versus capaciteit op verzinkdatum**. Bevestigde
  orders en reserveringen zijn gestapeld. Onderaan staat per dag het
  gemiddelde kg/stuk. Het bijschrift noemt welk percentage van de stuks
  geschat is.
- Dagoverzicht: kolommen `Stuks`, `Stuks_geschat`, `Capaciteit_stuks`,
  `Benutting_stuks_pct`, `Kg_per_stuk` en `Status_stuks`.
- Weekoverzicht: kolommen `Stuks`, `Capaciteit_stuks`, `Benutting_stuks_pct`
  en `Status_stuks`.
- Controletabel: kolommen `Stuks`, `Stuks_bron` (Orderregel / Geschat (klant
  kg/stuk) / Geschat (standaard kg/stuk)) en `Kg_per_stuk_gebruikt`.

## Nog niet aangepast

- De eerstvolgende leverdatum (advies) en het baliescherm rekenen nog alleen
  op kg.

## Code

- `shared.py`: `get_stuks_planning_actief`, `get_stuks_capaciteit_config`,
  `get_kg_per_stuk_default`, `stuks_uit_orderregels`,
  `bereken_kg_per_stuk_per_klant`, `koppel_stuks` en
  `validate_orderregels_xlsx`. `build_dashboard_data` heeft de nieuwe optionele
  parameter `capaciteit_stuks`.
- `viewer_app.py`: `make_stuks_matplotlib_chart`, de slider en de extra
  kolommen.
- `manager_app.py`: validatie van `Orderregels.xlsx`.
- `test_stuksplanning.py`: 8 nieuwe tests. De volledige suite slaagt (23).
