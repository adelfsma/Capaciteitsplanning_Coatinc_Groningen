# v2.5.14 — Productie dashboard iteratie 2: daily snapshots + OTIF-trend

## Overzicht

Iteratie 2 rondt het productie-dashboard af met historische data. Bij elke
viewer-load wordt automatisch een snapshot opgeslagen van de plan-tonnage en
OTIF-metrics voor de peildatum, zodat deze waarden in latere sessies accuraat
kunnen worden getoond.

- **TONNAGE PLAN historisch** — komt nu uit daily snapshots (i.p.v. leeg te
  blijven zoals in v2.5.13).
- **OTIF-trendgrafiek** — nieuwe sectie in het productie-dashboard tabblad.
- **Backfill-optie** — voor het geval de app een dag niet geopend is.

## Nieuw in `shared.py`

### Constants
```python
SNAPSHOT_HISTORY_FILE = "daily_snapshot_history.json"
SNAPSHOT_RETENTION_DAYS = 365
```

### Functies
- **`load_daily_snapshots() -> dict`** — laad snapshot-historie uit de bucket.
  Gecached met TTL 60s, automatisch geïnvalideerd na `save_daily_snapshot()`.
  Ontbreekt / corrupt → leeg dict (nooit crash).
- **`save_daily_snapshot(datum, tonnage_plan_kg, otif_result, overwrite=False,
  backfilled=False) -> bool`** — sla één entry op onder gegeven datum. Retourneert
  `True` als er iets is geschreven, `False` als er al een entry bestond en
  `overwrite=False`. Snoeit automatisch entries ouder dan
  `SNAPSHOT_RETENTION_DAYS`. `backfilled=True` markeert een handmatig
  toegevoegde snapshot.
- **`build_otif_trend_df(snapshots, days_back=90) -> pd.DataFrame`** — bouwt
  DataFrame met kolommen `Datum, OTIF_pct, Aantal_orders, Backfilled` voor de
  trend-grafiek. Filtert op de laatste N dagen en slaat entries zonder
  OTIF-percentage over (zodat de lijn niet naar 0 duikt).
- **`build_productie_dashboard_week(...)`** — nieuwe optionele parameter
  `snapshots: dict | None`. Voor historische dagen wordt TONNAGE PLAN nu
  gelezen uit de snapshot (was voorheen altijd leeg). Voor vandaag en verder
  blijft `dag_df` de bron.

### Snapshot-structuur (in `daily_snapshot_history.json`)
```json
{
  "2026-08-25": {
    "tonnage_plan_kg": 55000.0,
    "otif_pct": 85.0,
    "otif_totaal": 40,
    "otif_gereed": 34,
    "otif_niet_gereed": 6,
    "created_at": "2026-08-26T08:32:15",
    "backfilled": false
  }
}
```

## Nieuw in `viewer_app.py`

### Auto-snapshot bij load
Zodra de productie-dashboard tab de peildatum-OTIF berekend heeft, wordt er
automatisch een snapshot opgeslagen als er nog geen bestaat voor die datum.
`st.session_state` voorkomt dubbele save-attempts binnen één sessie.

De snapshot wordt **niet overschreven** als hij al bestaat: de eerste snapshot
van een dag (meest accurate) blijft leidend. Later herstellen kan via de
backfill-knop met `Overschrijf bestaand`.

### Backfill-sectie in beheer-expander
Onder het handmatige-invoer formulier is een backfill-sectie toegevoegd:
- Datum-picker (default: peildatum)
- Checkbox "Overschrijf bestaand"
- Knop "Backfill snapshot" — slaat een snapshot op onder de gekozen datum met
  de HUIDIGE plan- en OTIF-berekening, gemarkeerd als `backfilled: true`.

Bruikbaar voor dagen waarop de app niet werd geopend. Waarden zijn per
definitie minder accuraat dan on-the-day snapshots — dat wordt in de UI ook
zichtbaar (oranje punten in de trendgrafiek).

### Nieuwe grafiek: OTIF-trend
Onderaan het productie-dashboard, na de YTD-grafieken:
- Periode-selector: 30 / 60 / 90 / 180 / 365 dagen (default 90)
- Lijngrafiek met OTIF % per dag
- Twee horizontale streeplijnen: de bestaande OTIF-thresholds (groen/oranje)
  uit `OTIF_GAUGE_GREEN_THRESHOLD` en `OTIF_GAUGE_ORANGE_THRESHOLD`
- Backfilled snapshots worden **oranje** gerenderd (i.p.v. blauw) met een
  onderschrift dat aangeeft hoeveel punten dat zijn
- Y-as vast op 0-105% voor consistente visuele vergelijking

### Bestaande tab-tekst
Caption bovenaan is aangepast: "Historische plan-waarden komen beschikbaar in
een volgende versie" is nu verwijderd — die zijn er.

## Nieuwe / uitgebreide tests

`test_productie_dashboard.py`:
- **`TestSaveDailySnapshot`** — nieuwe datum, geen overschrijven zonder vlag,
  wel overschrijven met vlag, retentie snoeit oude entries.
- **`TestBuildOtifTrendDf`** — leeg input, entries zonder OTIF worden
  overgeslagen, filtering op days_back, sortering.
- **`TestBuildProductieDashboardWeekWithSnapshots`** — historisch uit snapshot,
  historisch zonder snapshot blijft leeg, toekomst blijft uit dag_df komen.

Totaal: **32 tests groen** (was 22 in v2.5.13).

## Bucket / deployment notities

### Nieuw bestand in de bucket
`daily_snapshot_history.json` wordt automatisch aangemaakt bij de eerste
succesvolle snapshot. Geen handmatige stap nodig, dezelfde bucket, dezelfde
credentials. Vanaf deploy groeit hij dagelijks; na 365 dagen is de omvang
stabiel (oude entries worden gesnoeid bij elke save).

### `manager_app.py`
**Ongewijzigd.** Alle nieuwe functionaliteit zit in de viewer.

### Verwacht storage-gebruik
Per snapshot ≈ 250 bytes. Bij 365 entries ≈ 90 KB. Verwaarloosbaar.

### Opbouw van historie
Iedere werkdag dat de app minimaal één keer wordt geopend, ontstaat een
on-the-day snapshot. Vanaf de installatie van v2.5.14 begint de historie op
te bouwen — de eerste dagen is de OTIF-trendgrafiek dus nog leeg / kort. Wil
je meteen historische punten? Gebruik dan de backfill-knop voor elke dag die
je wilt inhalen; die punten krijgen wel de oranje 'backfilled' vlag.

## Beperkingen die blijven

- **Backfill is een best-effort snapshot**: de huidige plan-berekening
  reconstrueert niet de exact-toen-geldende situatie. Als een historische dag
  wordt gebackfilled, staat er de nu-berekende plan-behoefte onder die datum.
  Voor de OTIF wordt de OTIF gebruikt van de peildatum die momenteel actief is
  in de sidebar. Dit is bewust: perfecte reconstructie zou een volledige
  historische orderportefeuille vereisen, wat niet beschikbaar is.
- **Één snapshot per dag**: de tweede open van de app op dezelfde dag
  overschrijft de eerste snapshot NIET. Rationale: de eerste snapshot van de
  dag (kort na daybreak) is doorgaans het meest bruikbaar omdat er dan nog
  weinig orders zijn afgehandeld.
