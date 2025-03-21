
# 📄 Formelstruktur: `=SUM()`

## 🔹 Syntax
`=SUM(Zahl1, [Zahl2], ...)`

### Parameter

| Parameter | Beschreibung                        | Pflichtangabe | erwarteter Datentyp         |
| --------- | ----------------------------------- | ------------- | --------------------------- |
| Zahl1     | Wert(e), die summiert werden sollen | Ja            | Zahl, Array, Bereich, Zelle |
| \[Zahl2]  | weitere(r) Wert(e)                  | Nein          | Zahl, Array, Bereich, Zelle |

## 🔍 In einfacher Sprache
*Was macht diese Funktion eigentlich – ohne Fachchinesisch?*
> Ganz simple Addition. Eine `SUM(1,2,3)` macht nichts anderes als ein `=1+2+3`. Interessant wird die `SUM`-Funktion, wenn es um das addieren von Arrays oder abstrakten Bereichen geht.


## 📌 Wofür ist das nützlich?
- Die `SUM`-Funktion findet immer dann Einsatz, wenn Werte miteinander summiert werden müssen.
- Wie schon erwähnt, besonders stark, wenn man beispielsweise ein gefiltertes Array addieren muss.

## 🔢 Beispiel
```excel
# Summiert das Array in A15#
=SUM(A15#) 

# Summiert das dynamisch generierte Array
=SUM(FILTER(A1:A150, A1:A150 > 10))

# Summiert einen Bereich
=SUM(A1:A150)

# Summiert eine Spalte einer definierten Tabelle
=SUM(tabTest[Umsatz])

# Summe zum Zählen aller Werte, die im Bereich A1:A150 über `5` sind (kürzere Alternative zu `COUNTIF`)
## Aufgepasst: Hier wird nicht der Bereich summiert, sondern das Vorkommen ausgegeben
=SUM(--(A1:A150>5))
```

*Doppelte Negierung im letzten Beispiel näher erläutert (kleinerer Bereich):
```
=SUM(--(A1:A5>5))
=SUM(--({10, 3, 10, 3, 10}>5))
=SUM(--({TRUE, FALSE, TRUE, FALSE, TRUE}))
=SUM({1,0,1,0,1})
=3
```

## 📊 Was kommt dabei raus?
| Eingabe-Daten                      | Ergebnis der Funktion   |
| ---------------------------------- | ----------------------- |
| A1:A10 gefüllt mit Zahlen und Text | Summe aller Zahlenwerte |

## 💡 Kreativer Einsatz
*Wie lässt sich die Funktion clever kombinieren oder zweckentfremden?*
- Wird oftmals in Dashboards oder Reports eingesetzt
- grundlegende Aggregatfunktion
- Kann mit Arrays, Bereichen, Tabellenspalten, etc. arbeiten
- Text wird als `0` interpretiert
	- `=SUM("text", 5, 5)` = `=SUM(0,5,5)`
- `SUM` als Bedingungsüberprüfung
	- Ursprung liegt in der booleschen Natur, da Excel alle Werte ungleich `0` als `1`/TRUE interpretiert

## ⚠ Typische Fehlerquellen
- Parameter liefert Fehlerwert
	- Die `SUM`-Funktion selbst liefert nur in den aller seltensten Fällen von sich aus einen Fehlerwert. Meistens resultiert ein Error aus einer fehlerhaften Formel innerhalb der `SUM`-Funktion.

