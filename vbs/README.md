# SDS_match_Callout — VBScript-Portierung

Dieser Ordner enthält eine **VBScript-Portierung** des Python-Tools aus dem
Hauptverzeichnis. Sie ist für Programmierer gedacht, die VBScript können, aber
**kein Python** — und die mit **regulären Ausdrücken (Regex) noch nicht vertraut**
sind. Deshalb erklärt diese README nicht nur *was* das Tool tut, sondern auch
*jeden einzelnen Regex* Zeichen für Zeichen.

Die gesamte Logik steckt in einer einzigen Datei: [`sds_match.vbs`](sds_match.vbs).

---

## Inhaltsverzeichnis

1. [Was macht das Tool?](#1-was-macht-das-tool)
2. [Voraussetzungen & Start](#2-voraussetzungen--start)
3. [Das Eingabeformat (Excel)](#3-das-eingabeformat-excel)
4. [Die Datenformate SDS und Callout](#4-die-datenformate-sds-und-callout)
5. [Konfiguration](#5-konfiguration)
6. [Regex-Crashkurs für Einsteiger](#6-regex-crashkurs-für-einsteiger)
7. [Jedes Suchmuster im Detail (Regex erklärt)](#7-jedes-suchmuster-im-detail-regex-erklärt)
8. [Normalisierung](#8-normalisierung)
9. [Scoring & Fallbildung (Union-Find)](#9-scoring--fallbildung-union-find)
10. [VBScript-Besonderheiten gegenüber Python](#10-vbscript-besonderheiten-gegenüber-python)
11. [Ein neues Suchmuster hinzufügen](#11-ein-neues-suchmuster-hinzufügen)
12. [Beispielausgabe](#12-beispielausgabe)

---

## 1. Was macht das Tool?

Eingelesen wird eine Excel-Datei mit Funkmeldungen. Jede Zeile ist entweder

- ein **Callout** (eine eingehende Alarmierung, beginnt mit `IncomingCallout:`) oder
- eine **SDS** (eine Statusmeldung / "Short Data Service"-Nachricht).

Ein Callout und eine oder mehrere SDS-Meldungen gehören oft zum **selben Einsatz**.
Das Tool versucht, diese Datensätze automatisch zu **Fällen** zusammenzufassen.
Dafür läuft folgende Verarbeitungskette (Pipeline):

```
Excel-Datei
   │
   ▼
[1] Einlesen        →  pro Zeile: Datum, Uhrzeit, Text
   │
   ▼
[2] Klassifizieren  →  "callout" oder "sds" (anhand des Präfix)
   │
   ▼
[3] Extrahieren     →  Regex ziehen Felder aus dem Text:
   │                    PLZ, Ort, Straße, Hausnummer, Objekt,
   │                    Schlagwort, (Stichwort), Einsatzmittel
   ▼
[4] Vergleichen     →  je zwei Datensätze: stimmen Felder überein?
   │                    pro Treffer gibt es Punkte (Gewichte)
   ▼
[5] Gruppieren      →  Datensätze mit genug Punkten UND im gleichen
   │                    Zeitfenster kommen in denselben "Fall"
   ▼
[6] Ausgeben        →  Fälle + Trefferstatistik
```

Im Python-Original ist diese Logik auf mehrere Dateien verteilt
(`sds_match.py`, `extractor.py`, `matcher.py`, `config.py`, `patterns/*.py`).
In der VBScript-Version steckt alles in **einer** Datei — die Abschnitte sind
durch große Kommentar-Banner (`' ====`) klar getrennt und tragen Hinweise,
welcher Python-Datei sie entsprechen.

---

## 2. Voraussetzungen & Start

- **Windows** mit installiertem **Microsoft Excel** (das Skript steuert Excel
  über COM-Automation fern, um die `.xlsx`-Datei zu lesen).
- Ausführung **immer mit `cscript`** (Konsole), nicht mit `wscript` (sonst
  öffnet jede Ausgabezeile ein eigenes Fenster). Das Skript bricht mit einem
  Hinweis ab, wenn es unter `wscript` läuft.

```bat
cscript //nologo sds_match.vbs C:\Pfad\zu\daten.xlsx
```

### Optionen

| Option                 | Bedeutung                                                          |
|------------------------|--------------------------------------------------------------------|
| `/sheet:<name|index>`  | Tabellenblatt wählen. **Index 0 = erstes Blatt** (Default).        |
| `/only:a,b,c`          | Nur diese Suchmuster aktivieren (kommagetrennt).                   |
| `/skip:a,b`            | Diese Suchmuster deaktivieren.                                     |
| `/list`                | Verfügbare Suchmuster auflisten und beenden.                       |
| `/show-content`        | Den Original-Text jeder Zeile mit ausgeben.                        |
| `/threshold:0.5`       | Score-Schwellwert (überschreibt den Default).                     |
| `/time-window:120`     | Zeitfenster in Minuten (überschreibt den Default).                |
| `/only-multi`          | Nur Fälle mit mindestens 2 Datensätzen anzeigen.                  |

Beispiel:

```bat
cscript //nologo sds_match.vbs daten.xlsx /only:plz,city,schlagwort /time-window:60 /only-multi
```

---

## 3. Das Eingabeformat (Excel)

Das Skript liest die **ersten drei Spalten** (A, B, C) ab **Zeile 1** (es gibt
**keine** Kopfzeile, die übersprungen wird):

| Spalte | Inhalt   | Beispiel                  |
|--------|----------|---------------------------|
| A      | Datum    | `06.05.2024`              |
| B      | Uhrzeit  | `12:34:56.789`            |
| C      | Text     | die eigentliche Meldung   |

Leere Textzellen (Spalte C) werden übersprungen. Datum und Uhrzeit werden zu
einem Zeitstempel zusammengesetzt. Kann eine Zeile nicht in ein Datum
umgewandelt werden, gilt sie als "ohne Zeit" und wird beim Zeitfenster
ignoriert (sie kann aber trotzdem über Feldübereinstimmungen einem Fall
zugeordnet werden).

> **Hinweis:** VBScript-Datumswerte kennen **keine Millisekunden**. Die `.789`
> aus dem Beispiel wird verworfen. Für ein Zeitfenster von z. B. 120 Minuten
> spielt das keine Rolle.

---

## 4. Die Datenformate SDS und Callout

Die Regex sind genau auf diese beiden Textformate zugeschnitten. Wer die
Formate kennt, versteht die Regex sofort.

### SDS (Statusmeldung)

Felder stehen in `||...||`-Blöcken mit Bezeichnern wie `EO:`, `OT:`, `OBJ:`, `SW:`:

```
... ||EO: 82515 Wolfratshausen; Margeritenstraße; 22a|| ... |+SW: #T2410#Rettung#...|+ ...
```

- `EO:` = Einsatzort: `PLZ Ort; Straße; Hausnummer`
- `OT:` = Ortsteil (optional, hat Vorrang vor dem Ort aus `EO:`)
- `OBJ:` = Objekt (optional, z. B. „Kirche")
- `SW:` = Schlagwort (zwischen `|+ ... |+`)

### Callout (Alarmierung)

Beginnt mit `IncomingCallout:`. Felder stehen ebenfalls zwischen doppelten
Pipes, aber in anderer Reihenfolge und ohne Bezeichner:

```
IncomingCallout:...||82515 Wolfratshausen - Wolfratshausen||Margeritenstraße 22a  og 2||Kirche||FL Sleh 11/1||#T2410#Rettung#Wohnung öffnen akut||...
```

- `||PLZ Ort - Gemeinde||` — Ort steht **links** vom ` - `
- direkt danach `||Straße Hausnummer  Zusatz||`
- optional ein `||Objekt||`-Block (z. B. „Kirche") — **darf keine Ziffern enthalten**
- danach **Einsatzmittel** (z. B. `FL Sleh 11/1`)
- das Schlagwort als `|#T2410#...|`-Block

---

## 5. Konfiguration

Im Python-Original lag die Konfiguration in `config.yaml`. VBScript hat keinen
YAML-Parser, deshalb stehen die Werte **als Konstanten oben in der Datei** und
die **Gewichte** in der Funktion `BuildExtractors`.

```vbscript
Const CALLOUT_PREFIX  = "IncomingCallout:"  ' so beginnt ein Callout
Const TIME_WINDOW_MIN = 120                 ' Zeitfenster in Minuten
Const SCORE_THRESHOLD = 0.5                 ' ab dieser Punktzahl: "zusammengehörig"
```

Die Gewichte (Punkte pro Feldtreffer) stehen in der letzten Spalte von
`BuildExtractors`:

| Muster          | Gewicht | zählt fürs Matching? |
|-----------------|---------|----------------------|
| `plz`           | 0.25    | ja                   |
| `street`        | 0.15    | ja                   |
| `house_number`  | 0.10    | ja                   |
| `city`          | 0.05    | ja                   |
| `objekt`        | 0.10    | ja                   |
| `schlagwort`    | 0.25    | ja                   |
| `stichwort`     | 0.10    | ja (matcht noch nichts) |
| `einsatzmittel` | 0.00    | **nein** (nur Anzeige) |

`einsatzmittel` hat `DisplayMode = "extra_line"` und zählt deshalb nie für den
Score — selbst wenn man ein Gewicht einträgt.

---

## 6. Regex-Crashkurs für Einsteiger

Ein **regulärer Ausdruck (Regex)** ist ein Suchmuster für Text. Statt „suche
genau dieses Wort" beschreibt man eine *Form*: „eine fünfstellige Zahl",
„ein Wort mit Großbuchstaben am Anfang" usw. VBScript benutzt dafür das
`RegExp`-Objekt.

### Die wichtigsten Bausteine

| Baustein   | Bedeutung                                                                 | Beispiel                     |
|------------|---------------------------------------------------------------------------|------------------------------|
| `abc`      | genau diese Zeichen                                                       | `EO` findet „EO"             |
| `.`        | **ein beliebiges** Zeichen                                                | `a.c` → „abc", „axc"         |
| `\d`       | eine Ziffer (0–9)                                                         | `\d\d` → „42"                |
| `\s`       | ein Leerraum-Zeichen (Space, Tab, Zeilenumbruch)                          |                              |
| `\w`       | ein „Wort-Zeichen" (Buchstabe, Ziffer, `_`)                               |                              |
| `[abc]`    | **eine** der Zeichen in der Klammer (eine *Zeichenklasse*)                | `[ab]` → „a" oder „b"        |
| `[a-z]`    | ein Zeichen aus dem Bereich a bis z                                       |                              |
| `[^|]`     | ein Zeichen, das **nicht** `|` ist (das `^` *negiert* die Klasse)         |                              |
| `\b`       | eine **Wortgrenze** (Übergang Buchstabe ↔ Nicht-Buchstabe)                | `\bstr\b` findet „str" allein |

### Wie oft? (Quantoren)

| Quantor   | Bedeutung                                  |
|-----------|--------------------------------------------|
| `?`       | 0 oder 1 mal (optional)                    |
| `*`       | 0 oder beliebig oft                        |
| `+`       | 1 oder beliebig oft                        |
| `{5}`     | genau 5 mal                                |
| `{0,2}`   | 0 bis 2 mal                                |

### Gierig vs. genügsam

- `+` und `*` sind standardmäßig **gierig**: sie nehmen so viel wie möglich.
- Ein angehängtes `?` macht sie **genügsam (lazy)**: so wenig wie möglich.
  - `[^|]+?` heißt: „so wenige Nicht-Pipe-Zeichen wie möglich" — wichtig, damit
    der Treffer beim **nächsten** `||` endet und nicht erst beim letzten.

### Gruppen

- `( ... )` ist eine **Fang-Gruppe (capture group)**: der Teil, der hier passt,
  wird gemerkt und kann ausgelesen werden. Das ist der **extrahierte Wert**.
- `(?: ... )` ist eine **Nicht-Fang-Gruppe**: dient nur dem Zusammenfassen
  (z. B. für einen Quantor), wird aber nicht gemerkt.
- `(?= ... )` ist ein **Lookahead**: „prüfe, ob hier Folgendes *steht*, aber
  verbrauche es nicht". Praktisch, um sich beim Suchen nur „umzuschauen".

### Zeichen, die man maskieren muss

Sonderzeichen verlieren ihre Bedeutung mit einem vorangestellten `\`:

| Maskiert | Bedeutet wörtlich    |
|----------|----------------------|
| `\|`     | ein Pipe-Zeichen `|` |
| `\.`     | ein Punkt `.`        |
| `\-`     | ein Minus `-`        |
| `\\`     | ein Backslash `\`    |

> **Wichtig für VBScript:** In einer Zeichenklasse `[...]` muss man Sonderzeichen
> meist *nicht* maskieren (z. B. ist `[.-]` der Punkt und der Bindestrich). Im
> Code ist trotzdem manchmal `[/\-]` geschrieben — das ist gleichbedeutend mit
> `[/-]` und nur zur Sicherheit/Lesbarkeit maskiert.

---

## 7. Jedes Suchmuster im Detail (Regex erklärt)

Alle Regex werden in der Funktion `InitRegex` angelegt. Über jedem steht im Code
die „logische" Form mit echten Umlauten. (Warum die Umlaute im Code über
`ChrW(...)` gebaut werden, steht in [Abschnitt 10](#10-vbscript-besonderheiten-gegenüber-python).)

### 7.1 Adresse — SDS (`plz`, `city`, `street`, `house_number`)

```
\|\|\s*EO:\s*(\d{5})\s+([^;|]+?)\s*;\s*([^;|]+?)\s*;\s*([^|]+?)\s*\|\|
```

Sucht den `||EO: ...||`-Block, z. B.
`||EO: 82515 Wolfratshausen; Margeritenstraße; 22a||`.

| Teil          | Bedeutung                                                                |
|---------------|--------------------------------------------------------------------------|
| `\|\|`        | zwei Pipes `||` (jedes `|` ist maskiert)                                  |
| `\s*`         | beliebig viele Leerzeichen                                               |
| `EO:`         | wörtlich „EO:"                                                            |
| `\s*`         | wieder optionale Leerzeichen                                            |
| `(\d{5})`     | **Gruppe 1 = PLZ**: genau fünf Ziffern                                  |
| `\s+`         | mindestens ein Leerzeichen                                              |
| `([^;|]+?)`   | **Gruppe 2 = Ort**: möglichst wenige Zeichen, die weder `;` noch `|` sind |
| `\s*;\s*`     | ein Semikolon, von Leerzeichen umgeben                                   |
| `([^;|]+?)`   | **Gruppe 3 = Straße**: wieder bis zum nächsten `;`                        |
| `\s*;\s*`     | nächstes Semikolon                                                       |
| `([^|]+?)`    | **Gruppe 4 = Hausnummer**: bis zum schließenden `||`                     |
| `\s*\|\|`     | optionale Leerzeichen, dann `||`                                         |

Dazu zwei optionale Zusatz-Blöcke:

```
\|\|\s*OT:\s*([^|]+?)\s*\|\|        ← Ortsteil (Gruppe 1), hat Vorrang vor dem EO:-Ort
\|\|\s*OBJ:\s*([^|]+?)\s*\|\|       ← Objekt (Gruppe 1)
```

Alle drei sind **`IgnoreCase`** (Groß-/Kleinschreibung egal).

### 7.2 Adresse — Callout (`plz`, `city`, `street`, `house_number`, `objekt`)

```
\|\|(\d{5})\s+([^|]+?)\|\|([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß.\-]*(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*?)\s+(HAUSNR)\b
```

Das ist der komplexeste Regex. Er sucht
`||82515 Wolfratshausen - Wolfratshausen||Margeritenstraße 22a ...`.

| Teil                                   | Bedeutung                                                                                 |
|----------------------------------------|-------------------------------------------------------------------------------------------|
| `\|\|`                                 | öffnendes `||`                                                                            |
| `(\d{5})`                              | **Gruppe 1 = PLZ** (fünf Ziffern)                                                         |
| `\s+`                                  | Leerzeichen                                                                               |
| `([^|]+?)`                             | **Gruppe 2 = „Ort - Gemeinde"** (alles bis zum nächsten `||`); der Ort wird später am ` - ` abgeschnitten |
| `\|\|`                                 | schließendes `||` des PLZ/Ort-Blocks                                                      |
| `([A-Za-zÄÖÜäöüß]`                     | **Gruppe 3 = Straße**: beginnt mit **einem Buchstaben** (inkl. Umlaute)                   |
| `[A-Za-zÄÖÜäöüß.\-]*`                  | …gefolgt von Buchstaben, Punkten und Bindestrichen                                        |
| `(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*?)`      | …optional weitere durch Leerzeichen/Bindestrich getrennte Wortteile (genügsam)            |
| `\s+`                                  | Leerzeichen zwischen Straße und Hausnummer                                                |
| `(HAUSNR)`                             | **Gruppe 4 = Hausnummer** (siehe unten)                                                   |
| `\b`                                   | Wortgrenze (die Hausnummer endet sauber)                                                  |

**Warum muss der Straßen-Block *direkt* auf den PLZ/Ort-Block folgen?** Damit
Einsatzmittel-Blöcke wie `||FL Sleh 11/1||` nicht versehentlich als Straße
erkannt werden.

Das eingebettete **Hausnummer-Muster** `HAUSNR`:

```
\d+\s?[a-zäöüß]{0,2}(?:\s?[/\-]\s?\d+\s?[a-zäöüß]{0,2})?
```

| Teil                          | Bedeutung                                                  | Beispiel        |
|-------------------------------|------------------------------------------------------------|-----------------|
| `\d+`                         | eine oder mehrere Ziffern                                  | `22`            |
| `\s?`                         | optionales Leerzeichen                                     |                 |
| `[a-zäöüß]{0,2}`              | 0 bis 2 Kleinbuchstaben                                    | `a` → `22a`     |
| `(?:\s?[/\-]\s?\d+\s?[a-zäöüß]{0,2})?` | optionaler zweiter Teil mit `/` oder `-`           | `22/24`, `5-7`  |

Danach wird **am Ende des Treffers** noch nach einem optionalen Objekt gesucht:

```
^[^|]*\|\|([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß .\-]*?)\|\|
```

- `^` verankert den Treffer am **Anfang des Reststrings** (alles ab dem Ende
  der Adresse). In Python war das `re.match(..., m.end())`; in VBScript schneiden
  wir den Reststring mit `Mid(content, m.FirstIndex + m.Length + 1)` heraus und
  verankern mit `^`.
- `[^|]*\|\|` überspringt den Zusatz nach der Hausnummer (z. B. „ og 2") bis zum
  nächsten `||`.
- **Gruppe 1 = Objekt**: ein Block, der mit einem Buchstaben beginnt und **keine
  Ziffern** enthält. So wird ein Objekt („Kirche") von einem Einsatzmittel
  („FL Sleh 11/1") unterschieden.

Falls die Volladresse nicht passt, gibt es einen **Fallback** nur für PLZ + Ort:

```
\|\|(\d{5})\s+([^|]+?)\|\|
```

### 7.3 Schlagwort (`schlagwort`)

Zwei verschiedene Regex — je nach Datentyp:

**Callout:**
```
\|(#[TBIR]\d{4}[a-zA-ZÖöÄäß#\- ]*)\|+
```

| Teil               | Bedeutung                                                          |
|--------------------|--------------------------------------------------------------------|
| `\|`               | ein Pipe                                                           |
| `(`                | Beginn **Gruppe 1 = Schlagwort**                                   |
| `#`                | wörtlich `#`                                                       |
| `[TBIR]`           | **genau einer** der Buchstaben T, B, I oder R                      |
| `\d{4}`            | genau vier Ziffern (z. B. `2410`)                                  |
| `[a-zA-ZÖöÄäß#\- ]*` | beliebig viele Buchstaben, Umlaute, `#`, `-` oder Leerzeichen     |
| `)`                | Ende der Gruppe                                                    |
| `\|+`              | ein oder mehrere abschließende Pipes                              |

Findet z. B. in `...|#T2410#Rettung#Wohnung öffnen akut|...` den Wert
`#T2410#Rettung#Wohnung öffnen akut`.

**SDS:**
```
\|+SW: (#[BTIR]\d{4}[a-zA-ZÖöÄäß#\- ]*)\|+
```

Gleicher Aufbau, aber mit dem Bezeichner `SW: ` davor und `|+` (ein oder mehrere
Pipes) an beiden Seiten.

### 7.4 Stichwort (`stichwort`) — Platzhalter

```
(?!x)x
```

Das ist ein **bewusst unmögliches** Muster: `(?!x)` sagt „an dieser Stelle darf
**kein** `x` folgen", direkt gefolgt von `x` — das kann nie gleichzeitig wahr
sein, also matcht es **nie**. Es ist ein Platzhalter, genau wie im Python-Original.
Trage hier später deine echten Stichwort-Regex ein (eine für Callout, eine für SDS).

### 7.5 Einsatzmittel (`einsatzmittel`) — nur Callout, nur Anzeige

Hier arbeiten zwei Regex zusammen.

**Alle `||...||`-Blöcke einsammeln:**
```
\|\|\s*([^|]+?)\s*(?=\|\|)
```

- `\|\|` öffnendes `||`
- `\s*([^|]+?)\s*` **Gruppe 1 = Inhalt** zwischen den Pipes (getrimmt)
- `(?=\|\|)` **Lookahead**: „danach muss `||` *stehen*", wird aber nicht
  mitverbraucht. Dadurch können sich aufeinanderfolgende Blöcke das trennende
  `||` *teilen* — sonst würde jeder zweite Block übersprungen. Dieser Regex
  läuft mit `Global = True`, liefert also **alle** Treffer.

Aus allen so gefundenen Blöcken werden anschließend die **Fremd-Blöcke**
herausgefiltert (im Code, nicht per Regex): der PLZ/Ort-Block, der Straßen-Block,
der Objekt-Block — und Schlagwort-Blöcke, die so erkannt werden:

```
^#[TBIR]\d{4}
```

(„beginnt mit `#`, einem von T/B/I/R und vier Ziffern"). Übrig bleiben die
echten Einsatzmittel, z. B. `FL Sleh 11/1`. Sie werden mit ` | ` verbunden
ausgegeben und zählen **nicht** für den Score.

---

## 8. Normalisierung

Bevor zwei Feldwerte verglichen werden, werden sie *normalisiert*, damit z. B.
„Margeritenstr." und „Margeritenstraße" als gleich gelten. Drei Regex
(`IgnoreCase`, `Global`):

| Regex          | Wirkung                                            |
|----------------|----------------------------------------------------|
| `\bstr\.?\b`   | „str" oder „str." als eigenes Wort → wird zu „straße" |
| `\bstrasse\b`  | „strasse" (ohne ß) → wird zu „straße"              |
| `\s+`          | mehrere Leerzeichen → ein einzelnes                |

Zusätzlich wird vorher klein geschrieben (`LCase`) und außen getrimmt (`Trim`).

---

## 9. Scoring & Fallbildung (Union-Find)

1. **Paarweiser Score** (`PairScore`): Für je zwei Datensätze wird über alle
   gewichteten Felder summiert. Stimmt der **normalisierte** Wert eines Feldes
   in beiden überein, kommt das Gewicht des Feldes dazu. Nur `inline`-Muster
   zählen (also nicht `einsatzmittel`).

2. **Zeitfenster**: Datensätze werden nach Zeit sortiert. Liegen zwei mehr als
   `TIME_WINDOW_MIN` Minuten auseinander, werden sie nicht mehr verglichen
   (Datensätze ohne Zeit werden ans Ende sortiert und nur über Felder gematcht).

3. **Schwellwert**: Erreicht ein Paar mindestens `SCORE_THRESHOLD` Punkte, gelten
   die beiden als zusammengehörig.

4. **Union-Find**: Ein klassischer Algorithmus, der zusammengehörige Elemente zu
   **Clustern** verschmilzt. Sind A↔B und B↔C verbunden, landen A, B und C im
   selben Fall — auch wenn A und C nie direkt verglichen wurden. Die Klasse
   `UnionFind` setzt das mit einem `parent`-Array um (`Find` sucht die Wurzel,
   `Union` hängt zwei Wurzeln aneinander).

5. **Fälle**: Jeder Cluster wird zu einem `CaseRec` mit `ScoreMax` (höchster
   Paar-Score im Fall) und `ScoreAvg` (Durchschnitt). Einzelne, unverbundene
   Datensätze sind „Singletons" (Fall mit nur einem Record).

---

## 10. VBScript-Besonderheiten gegenüber Python

Wer den Python-Code daneben legt, sollte diese Unterschiede kennen:

### Benannte Gruppen gibt es nicht
Python nutzt benannte Gruppen wie `(?P<plz>\d{5})` und liest sie mit
`m.group("plz")`. **VBScript kennt das nicht.** Hier sind die Gruppen
**nummeriert** und werden über `m.SubMatches(0)`, `m.SubMatches(1)`, …
ausgelesen (0-basiert!). In den Kommentaren und Tabellen oben steht jeweils,
welche Gruppennummer welches Feld ist.

### `SubMatches` ist 0-basiert
Die **erste** Fang-Gruppe ist `SubMatches(0)`, die zweite `SubMatches(1)` usw.
Ob überhaupt ein Treffer da ist, prüft man über `re.Execute(text).Count > 0`.

### Umlaute über `ChrW(...)`
Damit das Matching **unabhängig von der Datei-Kodierung** funktioniert, werden
die Umlaut-Zeichen in den Zeichenklassen nicht direkt hingeschrieben, sondern in
`InitGlobals` über ihren Unicode-Codepoint gebaut:

```vbscript
g_ae = ChrW(228)  ' ä      g_Ae = ChrW(196)  ' Ä
g_oe = ChrW(246)  ' ö      g_Oe = ChrW(214)  ' Ö
g_ue = ChrW(252)  ' ü      g_Ue = ChrW(220)  ' Ü
g_sz = ChrW(223)  ' ß
```

Eine Zeichenklasse wie `[A-Za-zÄÖÜäöüß]` entsteht dann per
String-Verkettung: `"[A-Za-z" & g_umlAll & "]"`. Im Code sieht das etwas sperrig
aus — die **logische** Form mit echten Umlauten steht aber immer im Kommentar
darüber, und diese README zeigt sie ebenfalls.

> Die deutschen Texte in der **Konsolenausgabe** sind hingegen direkt mit Umlaut
> geschrieben. Erscheinen sie als Buchstabensalat, ist das nur eine Anzeige-
> Frage der Datei-Kodierung (dann die `.vbs` als ANSI/Windows-1252 oder UTF-16
> speichern) — das **Ergebnis** ist davon nicht betroffen.

### Kein automatisches Laden der Muster
Python findet alle `patterns/*.py` automatisch (`pkgutil`). In VBScript werden
die Muster **explizit** in `BuildExtractors` registriert. Jedes Muster ist eine
Funktion (`ExtractPlz`, `ExtractCity`, …), die per `GetRef("…")` als
Funktionsreferenz im `Extractor`-Objekt hinterlegt und später mit
`fn(content, kind)` aufgerufen wird.

### Keine Millisekunden
VBScript-`Date` speichert nur bis zur Sekunde. Millisekunden aus der Excel-Zeit
werden verworfen (siehe Abschnitt 3).

### Performance
Der Vergleich ist von Natur aus „jeder mit jedem" (O(n²)), genau wie im Python-
Original. Auch die Sortierung ist ein einfacher Insertion-Sort. Für einige
Hundert bis wenige Tausend Zeilen ist das völlig ausreichend; bei sehr großen
Dateien wird es langsamer.

### YAML → Konstanten
Statt `config.yaml` (Python) stehen die Einstellungen als `Const` oben in der
Datei und die Gewichte in `BuildExtractors`.

---

## 11. Ein neues Suchmuster hinzufügen

Beispiel: ein neues Muster `funkrufname`.

1. **Regex anlegen** in `InitRegex` (Umlaute ggf. über die `g_uml…`-Bausteine):
   ```vbscript
   Set g_reFunk = NewRegex("Funk:\s*([A-Za-z0-9\-]+)", True, False)
   ```
   (vorher `Dim g_reFunk` zu den globalen Variablen hinzufügen)

2. **Extraktor-Funktion** schreiben (liefert String oder `""`):
   ```vbscript
   Function ExtractFunkrufname(content, kind)
       Dim ms : Set ms = g_reFunk.Execute(content)
       If ms.Count > 0 Then ExtractFunkrufname = ms(0).SubMatches(0) Else ExtractFunkrufname = ""
   End Function
   ```

3. **Registrieren** in `BuildExtractors` — Array vergrößern (`Dim arr(8)`) und
   eine Zeile ergänzen:
   ```vbscript
   Set arr(8) = MakeEx("funkrufname", "ExtractFunkrufname", "Funkrufname", 60, "inline", 0.10)
   ```
   Argumente: Name, Funktionsname (für `GetRef`), Beschreibung, Sortier-Reihenfolge,
   Anzeige-Modus (`"inline"` oder `"extra_line"`), Gewicht.

Fertig — das Muster taucht automatisch in `/list`, in der Ausgabe und in der
Statistik auf.

---

## 12. Beispielausgabe

```
Gelesen: 5 Datensätze (2 Callouts, 3 SDS)
Aktive Suchmuster: plz, street, house_number, city, objekt, schlagwort, stichwort, einsatzmittel
Zeitfenster: 120 min   Schwellwert: 0.50

Fall #1  3 Record(s) (1 Callout, 2 SDS)  score_max=0.65  score_avg=0.58
  [callout] 2024-05-06 12:34:56  plz='82515'  street='Margeritenstraße'  house_number='22a'  city='Wolfratshausen'  objekt='Kirche'  schlagwort='#T2410#Rettung'  stichwort='-'
           einsatzmittel: FL Sleh 11/1
  [sds    ] 2024-05-06 12:35:10  plz='82515'  street='Margeritenstraße'  house_number='22a'  city='Wolfratshausen'  objekt='-'  schlagwort='#T2410#Rettung'  stichwort='-'
  ...

Match-Statistik (pro Suchmuster):
  Muster               Callout             SDS          Gesamt
  plz                      2/2             3/3             5/5
  street                   2/2             3/3             5/5
  ...
```

Die Felder in der Record-Zeile sind die `inline`-Muster; `einsatzmittel` steht
als eingerückte Extra-Zeile darunter, weil es `DisplayMode = "extra_line"` hat.
