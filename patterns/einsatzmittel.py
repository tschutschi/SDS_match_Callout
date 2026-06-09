"""Einsatzmittel — nur in Callouts, mehrere Werte (Leerzeichen-getrennt).

display_mode='extra_line' → erscheint als separate Zeile unter dem Record und
fliesst NICHT in den Fall-Score ein, auch wenn ein Gewicht gesetzt wuerde.
"""

import re

from extractor import Extractor
from patterns._address import parse_address

# Einsatzmittel stehen als ||Wert||Wert||...||-Sequenz im Callout.
# Wir parsen NICHT die innere Struktur eines Einsatzmittels (das bricht, sobald
# ein Umlaut fehlt, z.B. 'TÖL' -> 'T L'), sondern ziehen robust alle Werte
# zwischen doppelten Pipes heraus und filtern dann die Fremd-Bloecke aus
# (Adresse + Schlagwort), die ebenfalls von '||' umschlossen sind.
# Lookahead (?=\|\|) erlaubt ueberlappende Matches durch '||'-Trennungen.
PIPE_BLOCKS_RE = re.compile(r"\|\|\s*([^|]+?)\s*(?=\|\|)")

# Schlagwort-Signatur (z.B. '#T2410#Rettung#...') — solche Bloecke sind
# keine Einsatzmittel und werden uebersprungen.
SCHLAGWORT_TOKEN_RE = re.compile(r"^#[TBIR]\d{4}")


def _extract(content: str, kind: str) -> str | None:
    if kind != "callout":
        return None
    # Die schon geparste Adresse liefert uns die Werte zum Aussortieren.
    addr = parse_address(content, kind)
    plz = addr["plz"]
    strasse = addr["strasse"]

    values: list[str] = []
    for raw in PIPE_BLOCKS_RE.findall(content):
        token = raw.strip()
        if not token:
            continue
        # PLZ-Ort-Block ueberspringen (z.B. '82515 Wolfratshausen - Wolfratshausen')
        if plz and token.startswith(plz):
            continue
        # Strassen-Block ueberspringen (z.B. 'Margeritenstraße 22a  og 2')
        if strasse and token.startswith(strasse):
            continue
        # Schlagwort-Block ueberspringen (z.B. '#T2410#Rettung#Wohnung öffnen akut')
        if SCHLAGWORT_TOKEN_RE.match(token):
            continue
        values.append(token)
    return " | ".join(values) if values else None


EXTRACTOR = Extractor(
    name="einsatzmittel",
    func=_extract,
    description="Einsatzmittel (nur Callout, mehrere durch Leerzeichen getrennt)",
    order=90,
    display_mode="extra_line",
)
