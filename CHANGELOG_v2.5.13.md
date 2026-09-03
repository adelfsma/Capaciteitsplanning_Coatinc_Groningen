# v2.5.13 — Statusaanpassing "Opgehangen" + Productie dashboard (iteratie 1)

## STATUS_MAP-aanpassing

- **`Opgehangen`** verschoven van `Niet verzinkt` → `Verzinkt`.
  - Rationale: opgehangen orders staan al aan het rek / op het punt door de bak te gaan en
    horen niet meer in de benodigde capaciteitsberekening mee te tellen.
  - `PC Opgehangen` bewust ongewijzigd op `Niet verzinkt`: die poedercoat-orders worden al
    op ordernummer-patroon (regex `^\d+C\d*$` of eindigend op `-C`) uit de
    capaciteitsplanning uitgesloten via `is_otif_excluded_coat_order()`.
- **OTIF-berekening niet geraakt**: `OTIF_STATUS_MAP` en `compute_otif()` gebruiken de rauwe
  `Status` (niet `Verzinkstatus`), dus de mapping `"opgehangen" → "Nee"` blijft actief.

## Nieuw: Productie dashboard (nieuwe viewer-tab)

Reproduceert het bestaande Excel-dashboard als extra tabblad in de viewer.

### Nieuwe bestanden in Supabase-bucket
- **`mis.xlsx`** (optioneel): cumulatieve MIS-export met kolommen `Datum, KG, ManuurTon,
  KG_Afkeur, Traverses, M2Trav`. Wordt dagelijks vervangen via de beheeromgeving.
- **`dashboard_manual.json`** (auto): handmatige velden per datum (klachten, storingstijd,
  veiligheidsincidenten). Structuur: `{"YYYY-MM-DD": {"klachten": N, "storingstijd_min": N,
  "veiligheidsincidenten": N}}`.

### Nieuwe [locatie]-secrets (allemaal optioneel)
```toml
[locatie]
# Scalar (statische norm)
norm_manuren_per_ton = 6.0
norm_traversen_per_dag = 75
norm_gem_gewicht_per_traverse = 1000

# OF gestructureerde periodes (winnen van scalar als aanwezig)
[[locatie.normen_manuren_per_ton]]
van = "2026-01-01"
tot = "2026-06-30"
waarde = 6.0

[[locatie.normen_manuren_per_ton]]
van = "2026-07-01"
tot = "9999-12-31"
waarde = 5.5
```
Fallback-volgorde: matchende periode → scalar → hardcoded default.

### Wijzigingen `shared.py`
- Import: `io` toegevoegd (voor MIS-parse uit bytes).
- Constants: `MIS_FILE`, `MIS_KOLOMMEN`, `DASHBOARD_MANUAL_FILE`, `PRODUCTIE_DASHBOARD_KOLOMMEN`, `_DAG_KORT_NL`.
- `MIS_FILE` toegevoegd aan `_BASE_OPTIONAL_FILES` + label in `OPTIONAL_FILE_LABELS`
  → verschijnt automatisch in de manager-UI zonder `manager_app.py` te wijzigen.
- Config-getters: `get_norm_manuren_per_ton(datum)`, `get_norm_traversen_per_dag(datum)`,
  `get_norm_gem_gewicht_per_traverse(datum)`. Interne helper `_resolve_norm_op_datum()`.
- Load-functies: `load_mis_data()`, `load_dashboard_manual()`.
- Save-functie: `save_dashboard_manual_entry(datum, klachten, storing, veiligheid)` —
  update per-datum in `dashboard_manual.json`, overige datums blijven ongewijzigd.
- Nieuwe validator: `validate_mis_xlsx(file_path)` (nog niet aangeroepen vanuit manager;
  MIS-load doet zelf een kolomcheck als tweede vangnet).
- Nieuwe build-functie: `build_productie_dashboard_week(mis_df, dag_df, manual_dict, jaar,
  weeknr, feestdagen)` — retourneert een DataFrame met 7 rijen (ma t/m zo).

### Wijzigingen `viewer_app.py`
- Nieuwe imports voor de bovenstaande functies.
- Bestaande tab "Dashboard" hernoemd naar **"Voorraad & OTIF"** (inhoud ongewijzigd).
- Nieuwe tab **"Productie dashboard"**:
  - Bovenaan: wachtwoord-beveiligde expander voor handmatige invoer per datum.
  - Week-selector (jaar + weeknr, default = huidige ISO-week).
  - Twee tabellen naast elkaar: gekozen week + vorige week.
  - Drie grafieken met normen-lijnen: manuren/ton, aantal traversen, gemiddeld gewicht.

### `manager_app.py`
- **Ongewijzigd.** MIS-upload verschijnt automatisch via `OPTIONAL_FILES`.
  Kolomvalidatie gebeurt bij load in de viewer (via `load_mis_data()`).

### Beperkingen iteratie 1
- **TONNAGE PLAN historisch**: leeg voor dagen vóór vandaag. Voor peildatum en toekomst
  wordt de actuele behoefte uit `build_dashboard_data()` gebruikt. Historische plan-waarden
  worden pas correct getoond na iteratie 2 (v2.5.14), die daily snapshots introduceert.
- **OTIF-historie**: nog niet gepersisteerd (komt in iteratie 2).

## Unit tests
- Nieuw bestand: `test_productie_dashboard.py` — 17 tests over:
  - Weekbouw (lege input, correcte ISO-week-datums, MIS-koppeling, weekend/feestdag,
    gewerkte zaterdag, TONNAGE PLAN alleen voor vandaag+).
  - Norm-resolver (default, scalar, periode-match, geen match, open einde).
  - Load-fallback bij ontbrekend/corrupt JSON-bestand.
- Bestaande `test_otif*.py` niet aangeraakt en blijven relevant.

## Deployment notities
- **Alle vestigingen**: geen actie vereist tenzij het productie-dashboard gebruikt wordt.
  Zonder `mis.xlsx` in de bucket toont de nieuwe tab enkel een waarschuwing.
- **CGR productie**: normen instellen in secrets voordat het dashboard live gaat.
- **Test eerst deploy**: verifieer op de test-app dat de manager-UI het nieuwe upload-vak
  toont, upload een MIS-bestand, en controleer de tab.
