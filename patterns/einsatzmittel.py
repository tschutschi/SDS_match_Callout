"""Einsatzmittel — nur in Callouts, mehrere Werte (Leerzeichen-getrennt).

display_mode='extra_line' → erscheint als separate Zeile unter dem Record und
fliesst NICHT in den Fall-Score ein, auch wenn ein Gewicht gesetzt wuerde.
"""

import re

from extractor import Extractor

# Einsatzmittel stehen als ||Wert||Wert||...||-Block im Callout.
# Wir parsen NICHT die innere Struktur eines Einsatzmittels (das bricht, sobald
# ein Umlaut aus der Datenquelle fehlt, z.B. "TÖL" -> "T L"), sondern ziehen
# robust alles heraus, was zwischen zwei doppelten Pipes steht.
# Lookahead (?=\|\|) stellt sicher, dass nur '||'-umschlossene Werte greifen und
# kein nachfolgender Freitext faelschlich mitgenommen wird.
EINSATZMITTEL_CALLOUT_RE = re.compile(r"\|\|\s*([^|]+?)\s*(?=\|\|)")


def _extract(content: str, kind: str) -> str | None:
    if kind != "callout":
        return None
    values = [v.strip() for v in EINSATZMITTEL_CALLOUT_RE.findall(content) if v.strip()]
    return " | ".join(values) if values else None


EXTRACTOR = Extractor(
    name="einsatzmittel",
    func=_extract,
    description="Einsatzmittel (nur Callout, mehrere durch Leerzeichen getrennt)",
    order=90,
    display_mode="extra_line",
)
