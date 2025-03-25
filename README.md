
# =THINK() – When SUM just won’t cut it

Willkommen in der etwas anderen Welt der Excel-Formeln – hier wird nicht nur gerechnet, sondern gedacht.

Dieses Repository ist kein Lexikon.  
Es ist ein Architekturprojekt für Denkstrukturen, Klartext und Transferintelligenz.

### Für alle:
- die Excel nicht nur bedienen, sondern verstehen wollen.  
- denen `=SUM()` zu wenig, aber Power BI zu viel ist.  
- die wissen, dass `LAMBDA` mehr kann als griechisch klingen.  
- die lieber **Systeme bauen** statt **Zellen füllen**.

---

### Was dich erwartet:
- **Strukturierte Lernpfade statt alphabetischer Funktionsstreuung**  
- **Praxisnahe Denkmodelle statt akademischer Syntaxfragmente**  
- **Erklärungen in Klartext – ohne technische Weichspülung**  
- **Anwendungsorientierte Mini-Architekturen statt Formelsalat**  
- Ein wenig **Stil zwischen Funktion, Form und Verstand**

Egal, ob du gerade erst Excel geöffnet hast oder bereits mit `CUBESET()` sprichst –  
hier findest du deinen Pfad zum systemischen Excel-Denken.

Bitte beachte, dass ich in diesem Repo mit den englischen Varianten der Formeln arbeite, da diese im internationalen Umfeld geläufig sind. Ebenfalls möchte ich hervorheben, dass die englische Variante Formeln zu schreiben anstelle eines Semikolons (;) ein normales Komma (,) verwendet. Also. Wenn du in den aufgezeigten Formeln ein Komma zur Parameterabtrennung siehst, tausche es in deinem Kopf einfach gegen ein Semikolon aus, wenn du mit der deutschen Version arbeitest.

---

### Aber warum das Ganze?

Über die Jahre habe ich beobachtet, dass vielen Anwendern der Umgang mit Syntax schwerfällt – insbesondere, diese nicht nur zu lesen, sondern auch sicher und nachhaltig umzusetzen.  
Experten hingegen beherrschen die Syntax meist problemlos, tun sich aber schwer damit, sie für „Normal-Sterbliche“ greifbar zu machen.

Dieses Repository will genau diese Lücke schließen:  
Eine offene Lernquelle, die sich mit den wichtigsten Funktionen beschäftigt, deren Anwendung verständlich macht und zeigt, dass Excel weit mehr ist als nur ein Listenpunkt im Bewerbungsabschnitt „Sonstiges“.

Ziel ist eine strukturierte Denkweise, die die Arbeit mit Excel nicht nur einfacher, sondern auch logischer und nachhaltiger macht.

> **Denn Excel ist keine Tabellenkalkulation. Excel ist ein Framework für strukturierte Informationsarchitektur.**

---

## 📚 Inhaltsstruktur – Lernstufen & Formeldenkmuster
### ⬜ 0. Aller Anfang ist normal
- Normalisierung von Daten
- Datentypen vs. Zellformatierung
- Tabellen vs. Zellen
- Listen und Arrays
- Namensmanager
- Best Practice für Lösungsansätze

### 🟩 1. Ich weiß, wie man Excel öffnet
_Einstieg in die grundlegende Funktionslogik_
- Erste Denkmodelle
- Logik vor Funktion verstehen

#### Formeln in diesem Kapitel:
| Funktion | Beschreibung | Typischer Nutzen |
|---------|----------------|------------------|
| `=1+1` | Rechnen im Zellenkontext | Einstieg in Zellarithmetik |
| `SUM()` | Addiert Zellbereiche | Basis jeder Berechnung |
| `AVERAGE()`, `MIN()`, `MAX()` | Grundlegende Statistikfunktionen | Einfache Auswertungen |
| `TODAY()`, `NOW()` | Datum/Zeit-Funktionen | Zeitstempel, dynamische Bezüge |
| `TEXT()` | Zahlen-/Datumsformatierung per Formel | Steuerung von Anzeigeformaten |
| `IF()` | Einfache Bedingungslogik | Entscheidungsstruktur auf Einzelebene |

### 🟨 2. SUMME kann ich
_Basisformeln effizient und strukturiert anwenden_
- Einstieg in strukturiertes Arbeiten
- Kontrollmechanismen, dynamische Bezüge, erste bedingte Aggregation

#### Formeln in diesem Kapitel:
| Funktion | Beschreibung | Typischer Nutzen |
|---------|----------------|------------------|
| `COUNTIF()`, `SUMIF()` | Bedingte Aggregation | Häufige Analyseaufgaben |
| `VLOOKUP()`, `HLOOKUP()` | Klassische Lookup-Mechanismen | Datenbezug aus einfachen Tabellen |
| `TEXTJOIN()` | Komplexere Stringkombination mit Trennzeichen | Reporting-Vereinfachung |
| `MATCH()`, `INDEX()` | Lookup-Kombination ohne SVERWEIS | Leistungsfähiger als klassische Verweise |

### 🟧 3. VERWEIS hab ich schonmal gehört
_Von klassischem Excel zur dynamischen Logik_
- Indexierungslogik
- Lookup-Kombinationen
- Strukturierte Referenzierung

#### Funktionen in diesem Kapitel:
| Funktion | Beschreibung | Typischer Nutzen |
|---------|----------------|------------------|
| `XLOOKUP()`, `XMATCH()` | Moderne Lookup-Funktion | Flexibel, fehlertolerant |
| `FILTER()`, `SORT()`, `SORTBY()` | Dynamische Ergebnisbereiche | Berichtsautomatisierung |
| `UNIQUE()` | Duplikatbereinigung | Dimensionstabellenaufbau |
| `SEQUENCE()` | Generiert Zahlen-/Indexreihen | Automatisierung |
| `CHOOSE()` | Szenariosteuerung | Modularisierung |
| `IFS()`, `SWITCH()` | Mehrstufige Bedingungslogik | Kompaktere Logikabfragen |

### 🟦 4. LET me introduce you to...
_Einstieg in strukturierte Architektur und Funktionskomposition_
- Alles wird einfacher durch LET
- Variable Strukturierung
- Lambda-Funktionen
- Dynamische Modularisierung

#### Formeln in diesem Kapitel:
| Funktion | Beschreibung | Typischer Nutzen |
|---------|----------------|------------------|
| `LET()` | Variablenstruktur in Excel | Rechenoptimierung, Lesbarkeit |
| `LAMBDA()` | Parametrisierte Funktionen | Wiederverwendbare Module |
| `BYROW()`, `BYCOL()` | Iterative Transformation | Zeilen-/Spaltenlogik ohne Hilfszellen |
| `TEXTSPLIT()`, `TEXTBEFORE()`, `TEXTAFTER()` | Textzerlegung | Datenaufbereitung |
| `TEXTJOIN()` + Arrays | Textkonsolidierung | Komplexe Ausgabeformate |

### 🟥 5. LAMBDA und die Welt der CUBEs…
_Architekturbasierte Formelentwicklung auf Expertenniveau_
- Funktionale Modellierung
- Komplexe Datenmodelle & CUBE-Funktionen
- Aggregationssteuerung auf Architektenniveau

#### Formeln in diesem Kapitel:
| Funktion | Beschreibung | Typischer Nutzen |
|---------|----------------|------------------|
| `LAMBDA()` inkl. Rekursion | Funktionen in Excel erstellen | Wiederverwendbare Bausteine |
| `MAP()`, `REDUCE()`, `SCAN()` | Iterative/akkumulative Logik | Listenverarbeitung auf funktionaler Ebene |
| `MAKEARRAY()` | dynamische Generierung | Generator für Matrix-Logik |
| `VSTACK()`, `HSTACK()` | Datenstrukturierung | Tabellenformate modular kombinieren |
| `CUBEMEMBER()` | Einzelwert aus Datenmodell | Dimensionselemente dynamisch abrufen |
| `CUBEVALUE()` | KPI-Abruf aus Datenmodell | Faktlogik in Berichten |
| `CUBEMEMBERPROPERTY()` | Zusatzinfo zu Member | Kontextuelle Ergänzung |
| `CUBERANKEDMEMBER()` | Rangfolge in Dimensionen | Top-N-Auswertungen |
| `CUBESET()` | Setbildung | Flexible Gruppierung |
| `CUBESETCOUNT()` | Mengenzählung in Sets | Steuerung/Aggregation |
| `CUBEKPIMEMBER()` | KPI aus Modell abrufen | Berichtskontextsteuerung |

---

## 🧭 Der Weg danach: PowerPivot & strukturierte Berichterstattung
- Einführung in Datenmodelle und warum sie Sinn machen
- Klartextbrücke zwischen Excel-Formeln und DAX
- Strukturierter Einstieg in PowerPivot
- Aggregationslogik und Measures verstehen
- Vom Sheet zur Reporting-Architektur

---

## 🧠 Zielgruppe
- Fortgeschrittene Excel-Anwender
- Modellierer, Berichtsarchitekten, BI-Übergangsdenker
- Technisch orientierte Kolleg:innen mit Interesse an Klartext-Systematik

---

## ⚙ Strukturprinzip
- **Markdown-basiert**
- **Modular pro Funktion oder Denkmodell**
- **Erweiterbar durch Tags und Querverlinkungen**
- **Keine Overhead-Syntax, sondern Praxisorientierung**

---

## 🧰 Work With Me
Wenn Sie Interesse an strukturierter Excel-Architektur, Reporting-Standardisierung oder Wissensvermittlung in Ihrem Unternehmen haben,  
kontaktieren Sie mich gerne: **[deine-mail@domain.xyz]**

Ich bin auch freiberuflich tätig – mit Fokus auf:
- Strukturierte Excel-Systeme
- PowerPivot- & DAX-Modellierung
- Reporting-Architektur
- Schulungsformate in Klartextsprache

→ Dieses Repository zeigt meine Denkweise. Der Transfer in Ihre Organisation ist mein Angebot.

---

## 📄 Lizenz & Beitrag
> _[Lizenztyp hier eintragen]_  
> Beiträge willkommen – Denkmodelle statt Codezeilen gefragt.
