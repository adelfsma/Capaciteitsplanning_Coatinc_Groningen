Capaciteitsplanning Coatinc Groningen v2.4.0

Toegevoegd:
- Segment/type materiaal per klant op basis van debtor-export.xlsx.
- Tweede dashboardgrafiek: tonnage per materiaaltype op verzinkdatum.
- Kolommen Segment_debtor_export en Materiaaltype in de controletabel 'Gebruikte gegevens'.
- Filter op Materiaaltype in de controletabel.
- Dagoverzicht uitgebreid met KG_Constructie, KG_Maatwerk, KG_Seriewerk en KG_Overig_onbekend.

Belangrijk:
- debtor-export.xlsx is optioneel om de app niet te blokkeren als het bestand nog niet is gepubliceerd.
- Als debtor-export.xlsx ontbreekt, worden regels getoond als 'Overig / onbekend'.
- Upload de debiteurenexport via de beheeromgeving in het veld 'debtor-export.xlsx'. De bestandsnaam in de cloud wordt dan automatisch debtor-export.xlsx.

Capaciteitsplanning Coatinc Groningen v2.1

Toegevoegd:
- wachtwoordbeveiliging in sidebar
- duidelijkere tekst voor laatste update

Capaciteitsplanning Coatinc Groningen v2.0

Opzet:
- viewer_app.py  -> kijk-app voor alle gebruikers
- manager_app.py -> beheer-app voor upload en publicatie
- shared.py      -> gedeelde logica
- published_data/ -> centrale gepubliceerde dataset
- templates/feestdagen_template.xlsx -> voorbeeldbestand voor feestdagen

Benodigde bestanden voor publicatie:
- OrderExport2G.xlsx
- Export-1.xlsx
- Export.xlsx
- Export+1.xlsx
- Export+2.xlsx
- Export+3.xlsx
- Export+4.xlsx
- feestdagen.xlsx

Optionele bestanden:
- Export_CGS.xlsx
- debtor-export.xlsx


Aanvulling v2.4.2
- TEST-markering is nu conditioneel gemaakt via Streamlit Secrets.
- Dezelfde code kan veilig op Test en Main draaien.
- Alleen apps met [app] environment = "test" tonen de grote TEST-banner, [TEST] in de browser-tab en de extra waarschuwing in beheer.

Benodigde Secrets per omgeving:

Voor Test-viewer en Test-beheer:
[app]
environment = "test"

[supabase]
url = "..."
key = "..."
bucket = "..."

Voor Main-viewer en Main-beheer:
[app]
environment = "production"

[supabase]
url = "..."
key = "..."
bucket = "..."

Als [app] environment ontbreekt, wordt de app behandeld als productie en wordt geen TEST-banner getoond.
