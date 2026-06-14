# SDS_match_Callout

> **Hinweis – VBScript-Portierung:** Im Ordner [`vbs/`](vbs/) liegt eine
> eigenständige Portierung dieses Tools nach **VBScript** (für Windows + Excel,
> ohne Python). Sie ist für VBS-Programmierer gedacht und enthält eine sehr
> ausführliche [`vbs/README.md`](vbs/README.md), die u. a. **jeden einzelnen
> Regex Zeichen für Zeichen** erklärt. Wer in regulären Ausdrücken einsteigen
> möchte, findet dort den besten Startpunkt.

Tool zum Einlesen von Funkmeldungen aus einer Excel-Datei, zum Erkennen
strukturierter Felder per Suchmuster (Regex) und zum Zusammenfassen
zusammengehöriger Datensätze zu **Fällen** (gewichteter Score + Zeitfenster).

---

## Inhalt

1. [Überblick](#überblick)
2. [Installation](#installation)
3. [Aufruf & Optionen](#aufruf--optionen)
4. [Eingabeformat (Excel)](#eingabeformat-excel)
5. [Datenformate: SDS und Callout](#datenformate-sds-und-callout)
6. [Projektstruktur](#projektstruktur)
7. [Konfiguration (`config.yaml`)](#konfiguration-configyaml)
8. [Die Suchmuster (Extraktoren)](#die-suchmuster-extraktoren)
9. [Scoring & Fallbildung](#scoring--fallbildung)
10. [Ein eigenes Suchmuster hinzufügen](#ein-eigenes-suchmuster-hinzufügen)
11. [Beispielausgabe](#beispielausgabe)

---

## Überblick

Eingelesen wird eine Excel-Datei mit Funkmeldungen. Jede Zeile ist entweder

- ein **Callout** (eingehende Alarmierung; Text beginnt mit `IncomingCallout:`) oder
- eine **SDS** (Statusmeldung / „Short Data Service").

Ein Callout und die dazugehörigen SDS-Meldungen gehören oft zum **selben
Einsatz**. Das Tool gruppiert solche Datensätze automatisch. Verarbeitungskette:

```
Excel-Datei
   │
   ▼
[1] Einlesen        →  pro Zeile: Datum, Uhrzeit, Text          (sds_match.py)
   │
   ▼
[2] Klassifizieren  →  "callout" oder "sds" anhand des Präfix   (extractor.py)
   │
   ▼
[3] Extrahieren     →  Regex ziehen Felder aus dem Text:        (patterns/*.py)
   │                    plz, city, street, house_number, objekt,
   │                    schlagwort, stichwort, einsatzmittel
   ▼
[4] Vergleichen     →  je zwei Records: übereinstimmende Felder (matcher.py)
   │                    geben gewichtete Punkte (pair_score)
   ▼
[5] Gruppieren      →  Records über Schwellwert + Zeitfenster   (matcher.py)
   │                    werden per Union-Find zu Fällen verschmolzen
   ▼
[6] Ausgeben        →  Fälle + Trefferstatistik                 (sds_match.py)
```

---

## Installation

Voraussetzung: **Python 3.10+** (der Code nutzt `X | None`-Typannotationen).

```bash
pip install -r requirements.txt
```

Abhängigkeiten (`requirements.txt`):

- `pandas>=2.0` – Excel-Einlesen / DataFrame
- `openpyxl>=3.1` – `.xlsx`-Engine für pandas
- `pyyaml>=6.0` – `config.yaml` laden

---

## Aufruf & Optionen

```bash
python sds_match.py <pfad/zur/datei.xlsx> [Optionen]
```

| Option                | Bedeutung                                                              |
|-----------------------|------------------------------------------------------------------------|
| `--sheet <name|idx>`  | Tabellenblatt (Name oder Index; Default `0` = erstes Blatt).           |
| `--config <pfad>`     | Pfad zur Config (Default `config.yaml`).                               |
| `--only a,b,c`        | Nur diese Suchmuster aktivieren (kommagetrennt).                       |
| `--skip a,b`          | Diese Suchmuster deaktivieren.                                         |
| `--list`              | Verfügbare Suchmuster auflisten und beenden.                           |
| `--show-content`      | Den Original-Text jedes Records mit ausgeben.                          |
| `--threshold <float>` | Score-Schwellwert (überschreibt die Config).                          |
| `--time-window <min>` | Zeitfenster in Minuten (überschreibt die Config).                     |
| `--only-multi`        | Nur Fälle mit ≥ 2 Records anzeigen.                                    |

Beispiele:

```bash
# Alle Muster, Standard-Config
python sds_match.py daten.xlsx

# Nur ausgewählte Muster, engeres Zeitfenster, nur echte Paare
python sds_match.py daten.xlsx --only plz,city,schlagwort --time-window 60 --only-multi

# Verfügbare Suchmuster anzeigen
python sds_match.py --list
```

---

## Eingabeformat (Excel)

Gelesen werden die **ersten drei Spalten** ab **Zeile 1** (keine Kopfzeile):

| Spalte | Inhalt   | Beispiel          |
|--------|----------|-------------------|
| A      | Datum    | `06.05.2024`      |
| B      | Uhrzeit  | `12:34:56.789`    |
| C      | Text     | die Meldung       |

- Leere Textzellen werden übersprungen.
- Datum + Uhrzeit werden zu einem `datetime` zusammengesetzt
  (`parse_timestamp`); unterstützte Formate:
  `%d.%m.%Y %H:%M:%S.%f` und `%d.%m.%Y %H:%M:%S`.
- Führende Apostrophe (`'`) und ein Komma als Dezimaltrenner in der Zeit werden
  toleriert.
- Lässt sich keine gültige Zeit bilden, gilt der Record als „ohne Zeit" – er
  wird beim Zeitfenster ans Ende sortiert, kann aber weiterhin über
  Feldübereinstimmungen einem Fall zugeordnet werden.

---

## Datenformate: SDS und Callout

Die Regex sind genau auf diese beiden Textformate zugeschnitten.

### SDS (Statusmeldung)

Felder in `||...||`-Blöcken mit Bezeichnern (`EO:`, `OT:`, `OBJ:`, `SW:`):

```
... ||EO: 82515 Wolfratshausen; Margeritenstraße; 22a|| ... |+SW: #T2410#Rettung#...|+ ...
```

- `EO:` Einsatzort als `PLZ Ort; Straße; Hausnummer`
- `OT:` Ortsteil (optional, hat Vorrang vor dem Ort aus `EO:`)
- `OBJ:` Objekt (optional)
- `SW:` Schlagwort (zwischen `|+ ... |+`)

### Callout (Alarmierung)

Beginnt mit `IncomingCallout:`. Felder zwischen doppelten Pipes, ohne Bezeichner:

```
IncomingCallout:...||82515 Wolfratshausen - Wolfratshausen||Margeritenstraße 22a  og 2||Kirche||FL Sleh 11/1||#T2410#Rettung#Wohnung öffnen akut||...
```

- `||PLZ Ort - Gemeinde||` – Ort steht links vom ` - `
- direkt danach `||Straße Hausnummer  Zusatz||`
- optional `||Objekt||` (darf **keine** Ziffern enthalten – grenzt es von
  Einsatzmitteln ab)
- danach Einsatzmittel (z. B. `FL Sleh 11/1`)
- Schlagwort als `|#T2410#...|`-Block

---

## Projektstruktur

| Datei                    | Aufgabe                                                                     |
|--------------------------|-----------------------------------------------------------------------------|
| `sds_match.py`           | CLI / Hauptprogramm: Excel lesen, Ausgabe, Argumente.                       |
| `config.py`              | Lädt `config.yaml` (mit Defaults, Deep-Merge) in ein `Config`-Dataclass.    |
| `extractor.py`           | `Extractor`-Framework, lädt `patterns/*.py` automatisch, klassifiziert, normalisiert. |
| `matcher.py`             | `Record`/`Case`, `pair_score`, Union-Find, `find_cases`.                    |
| `patterns/_address.py`   | Gemeinsamer, gecachter Adress-Parser für SDS und Callout.                   |
| `patterns/plz.py` u. a.  | Je ein Suchmuster (`EXTRACTOR`-Objekt) pro Feld.                            |
| `config.yaml`            | Präfix, Zeitfenster, Schwellwert, Gewichte.                                 |
| `vbs/`                   | **VBScript-Portierung** (siehe Hinweis oben).                              |

Das Extraktor-Framework lädt **alle** Muster automatisch: `load_patterns()`
durchsucht das Paket `patterns/` mit `pkgutil` und sammelt jede Modul-Variable
namens `EXTRACTOR` ein. Eine neue Datei `patterns/xyz.py` mit einem
`EXTRACTOR`-Objekt wird also ohne weitere Registrierung erkannt.

---

## Konfiguration (`config.yaml`)

```yaml
classification:
  callout_prefix: "IncomingCallout:"   # so beginnt ein Callout

matching:
  time_window_minutes: 120             # Zeitfenster
  score_threshold: 0.5                 # ab dieser Punktzahl: "zusammengehörig"

weights:                               # Gewichte je Suchmuster (Summe ≠ 1.0 erlaubt)
  plz: 0.25
  schlagwort: 0.25
  stichwort: 0.10
  street: 0.15
  city: 0.05
  house_number: 0.10
  objekt: 0.10
  # einsatzmittel: nur Anzeige — kein Gewicht (wird ohnehin ignoriert)
```

- `config.py` mergt die Datei über `DEFAULTS` (Deep-Merge), fehlende Werte
  werden also ergänzt.
- Muster **ohne** Gewichtseintrag zählen mit Gewicht 0 (also gar nicht) fürs
  Matching.
- `--threshold` und `--time-window` überschreiben die Config zur Laufzeit.

---

## Die Suchmuster (Extraktoren)

Jedes Muster ist ein `Extractor` (siehe `extractor.py`) mit:

- `name` – eindeutiger Name (für `--only`/`--skip`, Gewichte, Anzeige)
- `func(content, kind) -> str | None` – zieht den Wert heraus
- `order` – Sortierung in Liste/Ausgabe
- `normalize` – Wert-Normalisierung vor dem Vergleich (Default:
  `default_normalize` – lower, trim, `str.`/`strasse` → `straße`, Mehrfach-
  Leerzeichen zusammenfassen)
- `display_mode` – `"inline"` (zählt fürs Matching) oder `"extra_line"`
  (nur Anzeige, **nie** im Score)

| Name            | Order | Gewicht | Mode         | Beschreibung                                                  |
|-----------------|-------|---------|--------------|--------------------------------------------------------------|
| `plz`           | 10    | 0.25    | inline       | Postleitzahl                                                 |
| `street`        | 20    | 0.15    | inline       | Straße ohne Hausnummer                                       |
| `house_number`  | 25    | 0.10    | inline       | Hausnummer (Zahl + optional Buchstaben/Bruch)                |
| `city`          | 30    | 0.05    | inline       | Ort                                                          |
| `objekt`        | 35    | 0.10    | inline       | Objekt (z. B. Kirche, Schule)                                |
| `schlagwort`    | 40    | 0.25    | inline       | Schlagwort (`B#/T#/I#/R#` + 4 Ziffern)                       |
| `stichwort`     | 50    | 0.10    | inline       | **Platzhalter** – Regex matcht noch nichts (TODO)            |
| `einsatzmittel` | 90    | –       | extra_line   | Einsatzmittel (nur Callout, mehrere Werte)                   |

Die meisten Adressfelder (`plz`, `city`, `street`, `house_number`, `objekt`)
holen ihren Wert aus dem gemeinsamen, mit `@lru_cache` gepufferten
`parse_address()` in `patterns/_address.py` – so wird jede Zeile nur einmal
geparst.

> **Regex im Detail:** Eine Zeichen-für-Zeichen-Erklärung jedes einzelnen Regex
> (inkl. Hausnummer-, Schlagwort- und Einsatzmittel-Muster) steht in
> [`vbs/README.md`](vbs/README.md). Die Muster sind in beiden Projekten
> inhaltlich identisch.

---

## Scoring & Fallbildung

In `matcher.py`:

1. **`pair_score(a, b)`** summiert über alle gewichteten Felder: stimmt der
   **normalisierte** Wert eines Feldes in beiden Records überein, kommt das
   Gewicht hinzu. Nur `inline`-Muster zählen.
2. **Zeitfenster:** Records werden nach Zeit sortiert (Records ohne Zeit ans
   Ende). Liegen zwei mehr als `time_window_minutes` auseinander, wird der
   Vergleich abgebrochen.
3. **Schwellwert:** Erreicht ein Paar `score_threshold`, gelten beide als
   verbunden.
4. **Union-Find:** Verbundene Records werden zu Clustern verschmolzen (A↔B und
   B↔C ⇒ {A,B,C}).
5. **`Case`:** Jeder Cluster wird zu einem Fall mit `score_max` und `score_avg`.
   Einzelne, unverbundene Records sind „Singletons".

---

## Ein eigenes Suchmuster hinzufügen

Neue Datei `patterns/funkrufname.py`:

```python
import re
from extractor import Extractor

FUNK_RE = re.compile(r"Funk:\s*([A-Za-z0-9\-]+)")

def _extract(content: str, kind: str) -> str | None:
    m = FUNK_RE.search(content)
    return m.group(1) if m else None

EXTRACTOR = Extractor(
    name="funkrufname",
    func=_extract,
    description="Funkrufname",
    order=60,
)
```

Optional ein Gewicht in `config.yaml` ergänzen (`funkrufname: 0.10`). Das Muster
wird automatisch geladen und taucht in `--list`, Ausgabe und Statistik auf.

---

## Beispielausgabe

```
Gelesen: 5 Datensätze (2 Callouts, 3 SDS)
Aktive Suchmuster: plz, street, house_number, city, objekt, schlagwort, stichwort, einsatzmittel
Zeitfenster: 120 min   Schwellwert: 0.50
Fälle: 2 (1 mit mehreren Records, 1 Singletons)

Fall #1  3 Record(s) (1 Callout, 2 SDS)  score_max=0.65  score_avg=0.58
  [callout] 2024-05-06 12:34:56.000  plz='82515'  street='Margeritenstraße'  house_number='22a'  ...
           einsatzmittel: FL Sleh 11/1
  [sds    ] 2024-05-06 12:35:10.000  plz='82515'  street='Margeritenstraße'  house_number='22a'  ...

Match-Statistik (pro Suchmuster):
  Muster               Callout             SDS          Gesamt
  plz                      2/2             3/3             5/5
  ...
```

Die Felder in der Record-Zeile sind die `inline`-Muster; `einsatzmittel`
erscheint als eingerückte Extra-Zeile (`display_mode="extra_line"`).
