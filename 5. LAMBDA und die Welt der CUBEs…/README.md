# 5. LAMBDA und die Welt der CUBEs…

## Was erwartet dich in diesem Kapitel?

Willkommen im Maschinenraum von Excel – dort, wo Funktionslogik auf Datenmodellierung trifft. Wenn wir vorher schon schnell unterwegs waren, befinden wir uns jetzt in der Formel1 (höhö, Excelwitz...verstehst du...Formel...1...ach egal).

Was hier beginnt, ist **kein Formelkatalog** mehr. Es ist der Einstieg in formale Modellbildung, rekursive Verarbeitung, Datenstrukturierung und multidimensionale Auswertung.

Dieses Kapitel ist nicht dafür da, um einzelne Funktionen zu lernen. Es ist dafür da, Systeme zu bauen, denn HINTER Excel gibt es NOCH EIN Excel. Die mystische der Datenmodelle. Mit eigener Sprache (genannt DAX), eigener Logik, Measures und was weiß ich.
Das ist aber gar nicht das, worauf wir uns hier konzentrieren wollen. Mit den Cubefunktionen haben wir nämlich die Möglichkeit von unserem Workbook aus auf das Datenmodell zuzugreifen. Und umgehen damit sämtliche Restriktionen, die eventuell mit Pivottabellen einhergehen.
Cube-Funktionen sind – zumindest technisch – ganz normale Excel-Funktionen. Und genau deshalb lassen sie sich bis zu einem gewissen Grad mit Arrays und anderen Formelstrukturen kombinieren. Wir können Datenmodell iterativ durchlaufen und so Informationen automatisiert abrufen.
Hier wird uns vermutlich das ein oder andere mal der #GETTING_DATA Error über den Weg laufen.

### Funktionale Modellierung – LAMBDA 2.0  
Was als parametrisierte Funktion begann, wird hier zu einem logischen Framework.  
Mit `LAMBDA()` lassen sich jetzt rekursive Prozesse aufbauen, wiederverwendbare Funktionsmodule konstruieren und komplexe Logik kapseln.  
Das ist kein „Nice to have“, das ist der Punkt, an dem Excel beginnt, Programmstruktur zu imitieren – innerhalb von Zellen. An dieser Stelle beginnt der Moment, an dem der Kopf raucht – im besten Sinne.

### MAP, REDUCE, SCAN – Denken in Listen  
Mit `MAP()`, `REDUCE()` und `SCAN()` bricht Excel endgültig die Zellenlogik auf.  
Jetzt zählen nicht mehr einzelne Werte, sondern ganze Arrays – funktional verarbeitet, dynamisch transformiert.  
Hier beginnt die funktionale Transformation von Daten, ganz ohne VBA, ganz ohne Script – nur mit Formelsprache.
Am Anfang ein bisschen schwierig zu verstehen, aber genau deswegen mache ich das hier ja!

### MAKEARRAY, VSTACK, HSTACK – strukturelles Denken  
Wer `MAKEARRAY()` beherrscht, denkt **in Generatorlogik**. Wer `VSTACK()` und `HSTACK()` versteht, baut **dynamisch kombinierbare Tabellenstrukturen**. Da fällt mir auf, dass wir vermutlich noch auf `CHOOSECOLS()`und `CHOOSEROWS()` eingehen sollten... 
Das ist keine Datenmanipulation – das ist strukturelles Datenlayouting mit maximaler Steuerbarkeit.

### Einstieg in die multidimensionale Welt – CUBE-Funktionen  
Die `CUBE`-Formeln sind die Schnittstelle zur mehrdimensionalen Welt: Datenmodelle, KPIs, Dimensionen, Member, Sets, Measures, WTF.  
Hier arbeitest du nicht mehr mit Zellbezügen – **sondern mit analytischen Strukturen, wie sie in BI-Systemen Standard sind**.

Ob `CUBEMEMBER()` für dynamische Dimensionselemente, `CUBEVALUE()` für Faktlogik, `CUBESET()` für Gruppierungen oder `CUBERANKEDMEMBER()` für Top-N-Auswertungen – du sprichst hier die Sprache von **datenmodellbasiertem Reporting auf Systemebene**.

### Aggregationssteuerung auf Architektenniveau  
CUBE-Funktionen ermöglichen es, **Kennzahlen-Logik direkt in das Datenmodell zu legen** – präzise, filterbar, dynamisch aggregierbar.  
In Kombination mit den funktionalen Formeln entsteht ein **Berichtssystem, das skaliert – logisch, lesbar, wartbar**.

---

> Wer hier angekommen ist, baut keine Formeln mehr – er baut Systeme.

Dieses Kapitel ist kein Lernstoff – es ist der Bauplan für professionelle Excel-Modelle, **die WEIT über klassische Anwendungsfälle hinausgehen**.  
Und es ist der Punkt, an dem du mit Excel **nicht nur arbeitest, sondern denkst.**. Spätestens hier fängst du an deinen Tag mit `=COFFEE()` zu starten.
