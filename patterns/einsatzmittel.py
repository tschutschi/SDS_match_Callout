"""Einsatzmittel — nur in Callouts, mehrere Werte (Leerzeichen-getrennt).

display_mode='extra_line' → erscheint als separate Zeile unter dem Record und
fliesst NICHT in den Fall-Score ein, auch wenn ein Gewicht gesetzt wuerde.
"""

import re

from extractor import Extractor

# Erste Capture-Gruppe = der gesamte (Leerzeichen-getrennte) Wert-String.
EINSATZMITTEL_CALLOUT_RE = re.compile(r"(?!x)x")  # TODO: Callout-Einsatzmittel-Regex eintragen


def _extract(content: str, kind: str) -> str | None:
    if kind != "callout":
        return None
    m = EINSATZMITTEL_CALLOUT_RE.search(content)
    return m.group(1).strip() if m else None


EXTRACTOR = Extractor(
    name="einsatzmittel",
    func=_extract,
    description="Einsatzmittel (nur Callout, mehrere durch Leerzeichen getrennt)",
    order=90,
    display_mode="extra_line",
)
