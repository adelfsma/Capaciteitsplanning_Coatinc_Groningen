Capaciteitsplanning Coatinc Groningen v2.5.3

Gecombineerde grafiekversie:
- Schuine datumlabels op beide dashboardgrafieken.
- Aparte KG-per-traverse-instellingen voor Constructie (standaard 1100),
  Maatwerk (standaard 850) en Seriewerk (standaard 900).
- Rode lijn met het berekende aantal balken op de rechter y-as, inclusief
  waarde-labels per datum. Overig / onbekend gebruikt de Maatwerk-instelling.

OTIF depot-shift:
- Orders met "Aanleveren depot" = 1 (uit OrderExport2G.xlsx) hebben een
  deadline van 17:00 op hun leverdatum. Omdat de brondata rond 08:30
  wordt geëxporteerd, worden deze orders pas de eerstvolgende werkdag
  op OTIF getoetst.
- Op peildatum P worden meegeteld:
    (a) niet-depot orders met leverdatum == P
    (b) depot-orders met leverdatum == vorige werkdag (P)
- 'Vorige werkdag' slaat weekenden én feestdagen over.
- Backwards compatible: zonder de kolom "Aanleveren depot" of met de
  vlag apply_depot_shift=False valt de logica terug op het oude gedrag.

In de UI:
- Onder de OTIF-dashboardkaart komt een regel "waarvan N depot-order(s)
  van dd-mm-jjjj" als er depot-orders zijn meegeteld.
- Het OTIF-tabblad noemt de vorige-werkdag-datum en het aantal depot-
  orders in de caption boven de detailtabel.
- De detailtabel toont extra kolommen "Aanleveren depot" en "Is_depot"
  zodat de verschoven regels direct herkenbaar zijn.


Capaciteitsplanning Coatinc Groningen v2.5.2

Verbeteringen (compact + verfijnd):
- KPI-strook is nu compact zodat de focus op de capaciteitsgrafieken
  eronder blijft. De drie kaarten (Zwarte voorraad, Witte voorraad,
  OTIF) zijn met identieke inhoudsstructuur opgezet: label · groot
  getal · detailregel. Automatisch gelijke hoogte (128 px).
- De halfronde snelheidsmeter op het dashboard is vervangen door een
  slanke horizontale 3-zone bar met een klein driehoekje als indicator.
  Neemt vrijwel geen ruimte in en oogt professioneel.
- De volledige gauge (matplotlib, equal-zone verdeling) blijft
  beschikbaar in het OTIF-tabblad voor drill-down.


Capaciteitsplanning Coatinc Groningen v2.5.1

Verbeteringen:
- Uitlijning: de OTIF-tegel staat nu in dezelfde HTML-kaart als de
  voorraadkaarten. Alle drie de kaarten in het Dashboard hebben identieke
  hoogte, padding en border. De snelheidsmeter is als inline SVG in de
  kaart ingebed.
- Snelheidsmeter: nieuwe niet-lineaire schaal waarbij de rode, oranje
  en groene zone elk 60° van de arc beslaan. De groene zone (96–100 %)
  is nu duidelijk zichtbaar. Binnen elke zone is de naaldpositie
  lineair, dus goed afleesbaar.
- Labels op de meter tonen de zonegrenzen (0, 80, 96, 100), met
  op de matplotlib-variant in het OTIF-tabblad ook lichte tussenlabels.


Capaciteitsplanning Coatinc Groningen v2.5.0

Toegevoegd:
- KPI OTIF (On Time In Full): percentage orders op tijd gereed, zichtbaar
  als snelheidsmeter in het Dashboard-tabblad naast de voorraadkaarten.
- Peildatum voor OTIF is standaard de dag van de laatste publicatie
  (metadata.published_at); via de zijbalk aanpasbaar voor terugkijkende
  analyse.
- Nieuw tabblad "OTIF": kerncijfers, formule-uitleg, gefilterde detailtabel
  en een Excel-download in de stijl van "Gebruikte gegevens".
- Statusmapping conform Status_OTIF.xlsx (Ja / Nee / Nvt); reserveringen
  worden bij de OTIF-berekening niet meegeteld.
- In de testomgeving kan de OTIF-berekening óf tegen "Datum" óf tegen
  "Originele datum" worden getoetst (zodra de kolom in de weekexports
  aanwezig is). In productie blijft de businessdefinitie "Datum".

Drempels (aanpasbaar in shared.py):
- Groen  : ≥ 96,0 % op tijd
- Oranje : 80,0 % – 96,0 %
- Rood   : < 80,0 %

Formule:
  OTIF = (1 − aantal te laat / totaal aantal orders op peildatum) × 100 %

Bestanden ongewijzigd t.o.v. v2.4.3 (geen nieuwe uploads nodig): de OTIF
wordt afgeleid uit de reeds gebruikte weekexports (Export.xlsx). Optioneel
kunnen deze exports een "Originele datum"-kolom bevatten voor de tweede
variant.


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


Aanvulling v2.4.3
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


Wijzigingen v2.4.3:
- KPI-overzicht bovenin vervangen door een rustiger Voorraadoverzicht.
- Totaal aantal orders en totaal KG te verzinken verwijderd uit het bovenste dashboarddeel.
- Witte voorraad toegevoegd: som van gewichten met status Afgehaald, Nabewerking nog uitvoeren, PC Afgehaald en Coat gereed.
