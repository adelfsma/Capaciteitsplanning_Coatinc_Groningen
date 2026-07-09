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


Aanvulling v2.4.1-test
- Grote rode TEST-markering toegevoegd aan viewer_app.py en manager_app.py.
- Browser-tab titel begint met [TEST].
- Sidebar toont ook TESTOMGEVING.
- Manager toont een extra waarschuwing bij gedeelde bucket/secrets.

Let op bij promotie naar main:
- Deze TEST-markering staat bewust in de Test-branch.
- Zet APP_ENVIRONMENT in shared.py niet op TEST in productie/main.
