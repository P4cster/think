# Ok - was ist Syntax?
"Syntax" ist letztendlich nichts anderes, als ein hochgestochenes Wort für "Rezept". Quasi ein "Man nehme..." der IT-Welt.

Um Excel-Syntax, insbesondere jene auf der offiziellen Microsoft-Seite, zu verstehen, bedarf es ein bisschen Hintergrundwissen. Als Grundregel gilt:

- Jeder Parameter, der **NICHT** in eckigen Klammern steht ist eine Pflichtangabe. Ohne diese wird in jedem Fall ein Fehler ausgegeben.
- Jeder Parameter, der in eckigen Klammern steht ist eine optionale Angabe, hat aber einen Standardwert, auf den er zurück fällt, wenn man ihn nicht explizit angibt.

Nehmen wir als Beispiel eine `XLOOKUP`-Formel.

In der offiziellen Dokumentation wird diese mit der folgenden Syntax beschrieben:

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

Hier sehen wir nun drei Pflichtangaben und drei optionale Angaben.
- **lookup_value** -> Der Wert der gesucht werden soll
- **lookup_array** -> Der Bereich in dem lookup_value gesucht werden soll
- **return_array** -> Der Bereich aus dem der Wert in der gleichen Zeile zurück gegeben werden soll
- **\[if_not_found]** -> Was soll angezeigt werden, wenn der gesuchte Wert nicht gefunden wurde?
- **\[match_mode]** -> Wie soll auf Übereinstimmung überprüft werden?
- **\[search_mode]** -> Wie soll gesucht werden?

Was die einzelnen Parameter tun, sei mal dahingestellt. Wichtig ist nur zu verstehen, wie Excel diese Funktion jetzt in verschiedenen Ausführungen interpretiert.
Es ist zu erwähnen, dass in der Exceldokumentation jeder optionale Parameter in der Erklärung auch den Standard- bzw. Defaultwert benannt hat. 

Defaultwerte für unsere optionalen Parameter der `XLOOKUP`:
-** \[if_not_found]** -> \#N/A-Error
- **\[match_mode]** -> 0
- **\[search_mode]** -> 1

Um nun eine valide `XLOOKUP` zu schreiben, bedarf es also mindestens drei Parameter. Diejenigen, die nicht in eckigen Klammern stehen.

```
=XLOOKUP("test", A1:A300, C1:C300)
```

So wird das Wort "test" in der Spalte A gesucht und der Wert aus der gleichen Zeile aus Spalte C wird zurück gegeben. 

Excel selbst interpretiert diese Formel allerdings mit allen Parametern, die übergeben werden können - sofern nicht explizit angegeben mit den Defaultwerten. Die Formel die Excel nun eigentlich interpretiert sieht wie folgt aus:

```
=XLOOKUP("test", A1:A300, C1:C300, NA(), 0, 1)
```

Warum ist das wichtig? Von Zeit zu Zeit kann es sein, dass eine Formel nicht das ausgibt, was man erwartet und der "Fehler" darin liegt, dass man einen optionalen Parameter nicht passend angegeben hat.

Außerdem geben die optionalen Parameter durchaus mächtige Möglichkeiten frei. Im Falle der `XLOOKUP` ist das zum Beispiel die Möglichkeit eine Liste von unten nach oben zu durchsuchen, eine Binärsuche durchzuführen oder, insbesondere im Falle von Datumsuchen, die nächstbeste Übereinstimmung zu wählen. Sich mit den optionalen Parametern auseinanderzusetzen macht also in jedem Fall sehen und wenn es nur dafür ist, auf welche Standardwerte zurück gegriffen wird, wenn man sie nicht angibt.

So viel zum Syntaxverständnis.

Prinzipiell sind alle Parameter in der Dokumentation sehr deskriptiv gehalten, sofern man weiß was sie machen sollen.

Dazu muss gesagt sein, dass man weiß Gott nicht bei jeder Funktion wissen muss, wofür sie benötigt wird. Realistisch betrachtet, verwendet man vielleicht 10 - 15 % aller Formeln in alltäglichen Anwendungsfällen, wenn überhaupt. Es kommt eben nur drauf an, in welchem Feld man unterwegs ist. 

Abschließend: So gut wie jeder Parameter jeder Formel kann "extern" generiert werden. Hier entwickelt sich die Stärke von verschachtelten Funktionen. Wir haben diese Möglichkeit schon in der vorigen Lektion [[6. Best Practice für Lösungsansätze]] sehen können, in der sowohl die Reihen- als auch Spaltennummer der `INDEX`-Funktion durch `MATCH`-Funktionen generiert wurde.

Nochmal zur Übersicht:

```
=FUNCTION(required_param, [optional_param1], [optional_bool])
```

- **bool** ist dabei die erwartete Rückgabe von TRUE oder FALSE, oftmals auch als **Bedingung** bezeichnet.