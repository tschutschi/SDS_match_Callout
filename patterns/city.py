"""Ortsname — eigenständige Regex je kind, ohne Abhängigkeit von patterns.plz."""

import re

from extractor import Extractor

# Callout: PLZ steht direkt vor dem Ort (z.B. "82418 Murnau").
# Die PLZ-Form ist hier nur Anker — die Erkennung der PLZ als Suchmuster
# liegt weiterhin in patterns/plz.py.
CITY_CALLOUT_RE = re.compile(
    r"\b8(?:2|3)\d{3}\s+([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß.\-]*(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*)"
)

# SDS: kein PLZ-Kontext. Hier Deine SDS-spezifische Regex eintragen.
# Solange Platzhalter (matcht nie), gibt es fuer SDS-Zeilen keine Stadt.
CITY_SDS_RE = re.compile(r"(?!x)x")  # TODO: SDS-Ort-Regex eintragen


def _extract(content: str, kind: str) -> str | None:
    pattern = CITY_SDS_RE if kind == "sds" else CITY_CALLOUT_RE
    m = pattern.search(content)
    return m.group(1).strip() if m else None


EXTRACTOR = Extractor(
    name="city",
    func=_extract,
    description="Ortsname (Callout: nach PLZ, SDS: eigene Regex)",
    order=30,
)
