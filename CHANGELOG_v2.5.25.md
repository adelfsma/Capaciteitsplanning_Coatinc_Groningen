# v2.5.25 — Snellere viewer

## Aanleiding
De viewer voelde traag aan. Bij elke wijziging van een instelling of filter
werd de hele app opnieuw doorgerekend en getekend, inclusief alle tabbladen.
Daarnaast werden alle gepubliceerde Excel-bestanden elke 60 seconden opnieuw
gedownload en verwerkt.

## Gewijzigd
1. **De data blijft bewaard tot de volgende publicatie** (`shared.py`).
   - `load_published_data` gebruikt het tijdstip en de bestanden uit
     `metadata.json` als cachesleutel. Voorheen gebeurde de zware download en
     verwerking elke 60 seconden opnieuw.
   - Een nieuwe publicatie is nog steeds binnen ongeveer 60 seconden zichtbaar.
     De knop **Ververs data** werkt zoals voorheen.
   - Vangnet: na 6 uur, of op een nieuwe dag, wordt er altijd opnieuw geladen.
   - `load_metadata` is nu ook gecachet (60 s). Voorheen werd dit bestand bij
     élke klik gedownload. `save_metadata` leegt deze cache direct.
2. **Alleen het geopende tabblad wordt uitgevoerd** (`viewer_app.py`).
   - `st.tabs(..., on_change="rerun")`: de tabbladen zien er hetzelfde uit,
     maar alleen het geselecteerde tabblad rekent. Het baliescherm rekent dus
     ook niet meer mee als het niet open staat.
   - Wisselen van tabblad geeft nu een korte herberekening (0,1–0,8 s).
     Voorheen was dat direct, maar dan rekende elke klik ergens anders alle
     tabbladen.
   - Op een oudere Streamlit-versie zonder deze optie werkt de viewer zoals
     voorheen.
3. **Grafieken worden bewaard.** De kg-, stuks- en materiaalgrafiek worden pas
   opnieuw getekend als de dagdata verandert. Grafieken worden gedeeld tussen
   gebruikers met dezelfde instellingen. Resolutie 150 dpi (was 200), zonder
   zichtbaar verschil.
4. **Tijdmeting** in het tabblad Debug, onder "Snelheid": tijd per stap
   (data laden, planning, leverdatum, OTIF), het totaal van deze run en de
   duur van de vorige klik.

## Gemeten (export 08-10, 1.063 regels, 2 s gesimuleerde clouddownload)
| Actie | v2.5.24 | v2.5.25 |
|---|---|---|
| Instelling in zijbalk wijzigen (tabblad Voorraad & OTIF) | 2,5–2,6 s | 1,2–1,5 s |
| Zelfde, na meer dan 60 s niets doen | 4,7 s | 1,3 s |
| Andere tabbladen, na een wijziging | ~2,5 s (alles rekent) | 0,1–0,8 s |
| Terug naar een eerder gebruikte instelling | 2,5 s | < 0,5 s (cache) |

Op Streamlit Cloud is de winst na 60 s waarschijnlijk groter, omdat daar de
echte download uit Supabase plaatsvond. Controleer dat met het tabblad Debug.

## Niet gewijzigd
- De rekenlogica, de uitkomsten en `manager_app.py`.
- De lay-out en inhoud van alle tabbladen.

## Let op bij de review in GitHub Desktop
`viewer_app.py` toont veel gewijzigde regels. Dat komt doordat de inhoud van
elk tabblad één niveau verder is ingesprongen (`if _tab_open(...)`). Zet in
GitHub Desktop **Hide whitespace changes** aan om alleen de echte wijzigingen
te zien.

## Tests
- Alle 93 tests slagen.
- `test_otif.py` en `test_otif_e2e.py` zijn bijgewerkt naar de huidige
  `OTIF_STATUS_MAP`. Ze verwachtten nog de oude indeling, waarin
  "PC Afgehaald" Nvt was en "PC Opgehangen", "Nabewerking nog uitvoeren" en
  "meetrapport" Nee waren. Al die statussen tellen nu als op tijd (Ja).
- De OTIF-percentages in de tests veranderen niet: Nvt en Ja tellen allebei
  niet als te laat.
