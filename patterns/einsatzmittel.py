"""Einsatzmittel — nur in Callouts, mehrere Werte (Leerzeichen-getrennt).

display_mode='extra_line' → erscheint als separate Zeile unter dem Record und
fliesst NICHT in den Fall-Score ein, auch wenn ein Gewicht gesetzt wuerde.
"""

import re

from extractor import Extractor

# Eine Capture-Gruppe, '||' ausgeklammert, Alternativen nicht-fangend (?:...).
# findall liefert alle Einsatzmittel im Callout — sie werden zusammengefuehrt.
EINSATZMITTEL_CALLOUT_RE = re.compile(
    r"\|\|("
    r"(?:\w{2}\s\w+\s\d/?\d+/\d)"
    r"|(?:\w{2}\s\w+\s?\w+\s\w+\s\d/?\d?)"
    r")"
)


def _extract(content: str, kind: str) -> str | None:
    if kind != "callout":
        return None
    matches = EINSATZMITTEL_CALLOUT_RE.findall(content)
    # findall kann pro Treffer ein str oder ein Tupel (mehrere Gruppen) liefern.
    values = []
    for m in matches:
        val = next((g for g in m if g), "") if isinstance(m, tuple) else m
        val = val.strip()
        if val:
            values.append(val)
    return " | ".join(values) if values else None


EXTRACTOR = Extractor(
    name="einsatzmittel",
    func=_extract,
    description="Einsatzmittel (nur Callout, mehrere durch Leerzeichen getrennt)",
    order=90,
    display_mode="extra_line",
)
