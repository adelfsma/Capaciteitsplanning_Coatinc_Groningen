# v2.5.24 — Verzinkdatum poetsorders: leverdatum − 3 werkdagen

## Overzicht

Poetsorders kunnen een eigen verzinkdatum krijgen: **leverdatum min 3
werkdagen**, in plaats van min 2. Dat geeft tijd voor het poetsen tussen
verzinken en leveren. Depotorders blijven altijd op min 1, ook als het een
poetsorder is.

De regel staat standaard **uit**. Groningen blijft ongewijzigd.

## Aanzetten (Streamlit-secrets van viewer én beheer)

```toml
[locatie]
poetsen_verzink_werkdagen = 3
```

## Hoe een poetsorder wordt herkend

**Bevestigde deelorders (weekexports)**
1. `PrijsCategorie` bevat "poetsen" ("Poetsen licht ( B )" / "Poetsen zwaar ( A )").
2. Is `PrijsCategorie` leeg, dan geldt een terugval: het nummer eindigt op `VC<n>-V`.
3. Een `VC`-deelorder mét een `PrijsCategorie` zonder poetsen (coat zonder
   poetswerk) blijft op min 2.

**Reserveringen (OrderExport2G)**: een van deze kenmerken:
- `Gewicht poetsen` > 0
- `Leverdatum C` is gevuld
- een orderregel met `Subbonnummer V` = `…VC<n>-V` (alleen als `Orderregels.xlsx`
  gepubliceerd is)

Coat-delen (`-C`) worden niet gemarkeerd; die vallen buiten de planning.

## Verzinkdatum (volgorde)

1. Depot (`Aanleveren depot = 1`): leverdatum − 1 werkdag
2. Poetsorder: leverdatum − `poetsen_verzink_werkdagen`
3. Overig: leverdatum − 2 werkdagen

Reserveringen zonder `Leverdatum V` krijgen zoals voorheen de leverdatum
`Datum verzending` + 5 werkdagen. Bij een poetsreservering is de verzinkdatum
dan `Datum verzending` + 2 werkdagen.

## In de viewer

- Controletabel: nieuwe kolommen `Poetsorder` (Ja/Nee) en `Poetsorder_reden`.
- Het bijschrift onder de kg-grafiek noemt de regels voor depot en poetsen.
- De dag- en weekbelasting, het advies voor de eerstvolgende leverdatum en de
  balie gebruiken automatisch de nieuwe verzinkdatum.

## Bekende beperking

Bij reserveringen is poetswerk vaak niet herkenbaar. In de CAL-exports van
14-10-2026 was ongeveer 27% van de bevestigde verzinkregels een poetsorder,
tegen 3 van de 181 reserveringen. Niet-herkende poetsreserveringen blijven op
min 2 (bewuste keuze, geen schatting per klant).

## Code

- `shared.py`: `get_poetsen_verzink_werkdagen` en `markeer_poetsorders`.
  `build_dashboard_data` heeft de nieuwe optionele parameter `poetsen_offset`
  (None = waarde uit de secrets). Orderregels worden ook ingelezen als alleen
  de poetsregel actief is.
- `viewer_app.py`: de kolommen in de controletabel en het bijschrift.
- `test_poetsorders.py`: 5 nieuwe tests.
