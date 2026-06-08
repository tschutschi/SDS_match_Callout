"""Einsatzmittel — nur in Callouts, mehrere Werte (Leerzeichen-getrennt).

display_mode='extra_line' → erscheint als separate Zeile unter dem Record und
fliesst NICHT in den Fall-Score ein, auch wenn ein Gewicht gesetzt wuerde.
"""

import re

from extractor import Extractor
from patterns._address import parse_address

# Einsatzmittel stehen als ||Wert||Wert||...||-Sequenz im Callout.
# Wir parsen NICHT die innere Struktur eines Einsatzmittels (das bricht, sobald
<<<<<<< Updated upstream
# ein Umlaut fehlt, z.B. 'TÖL' -> 'T L'), sondern ziehen robust alle Werte
# zwischen doppelten Pipes heraus und filtern dann die Adress-Bloecke aus.
# Lookahead (?=\|\|) erlaubt ueberlappende Matches durch '||'-Trennungen.
PIPE_BLOCKS_RE = re.compile(r"\|\|\s*([^|]+?)\s*(?=\|\|)")
=======
# ein Umlaut aus der Datenquelle fehlt, z.B. "TÖL" -> "T L"), sondern ziehen
# robust alles heraus, was zwischen zwei doppelten Pipes steht.
# Lookahead (?=\|\|) stellt sicher, dass nur '||'-umschlossene Werte greifen und
# kein nachfolgender Freitext faelschlich mitgenommen wird.
EINSATZMITTEL_CALLOUT_RE = re.compile(r"\|\|(\w{2}\s\w+\s\d\/?\d+\/\d|\|\|\w{2}\s\w+\s?\w+\s\w+\s\d\/?\d?)")
>>>>>>> Stashed changes


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
        values.append(token)
    return " | ".join(values) if values else None


EXTRACTOR = Extractor(
    name="einsatzmittel",
    func=_extract,
    description="Einsatzmittel (nur Callout, mehrere durch Leerzeichen getrennt)",
    order=90,
    display_mode="extra_line",
)
