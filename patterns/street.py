"""Straße + Hausnummer — eigenständige Regex je kind, ohne patterns.plz-Abhängigkeit."""

import re

from extractor import Extractor

# Callout: PLZ steht typischerweise direkt nach der Straße ("Hauptstr. 5, 82418 ...").
# Die PLZ-Form ist hier nur Anker; das Erkennen der PLZ als eigenes Suchmuster
# liegt weiterhin in patterns/plz.py.
_PLZ_ANCHOR = re.compile(r"\b8(?:2|3)\d{3}\b")
STREET_CALLOUT_RE = re.compile(
    r"([A-Za-zÄÖÜäöüß.\-]+(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*\s+\d+\s*[a-zA-Z]?)\s*$"
)

# SDS: kein PLZ-Kontext. Hier Deine SDS-spezifische Strassen-Regex eintragen.
# Solange Platzhalter (matcht nie), gibt es fuer SDS-Zeilen keine Strasse.
STREET_SDS_RE = re.compile(r"(?!x)x")  # TODO: SDS-Strassen-Regex eintragen


def _extract(content: str, kind: str) -> str | None:
    text = " ".join(content.split())
    if kind == "callout":
        plz_match = _PLZ_ANCHOR.search(text)
        if not plz_match:
            return None
        before = text[: plz_match.start()].rstrip(" ,;:-")
        m = STREET_CALLOUT_RE.search(before)
        return m.group(1).strip() if m else None
    m = STREET_SDS_RE.search(text)
    return m.group(1).strip() if m else None


EXTRACTOR = Extractor(
    name="street",
    func=_extract,
    description="Straße + Hausnummer (Callout: vor der PLZ, SDS: eigene Regex)",
    order=20,
)
