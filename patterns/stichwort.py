"""Stichwort — getrennte Regex für Callout und SDS (Struktur wie schlagwort)."""

import re

from extractor import Extractor

# Beide Regex sind Platzhalter und matchen nichts.
# Erste Capture-Gruppe = der extrahierte Wert.
STICHWORT_CALLOUT_RE = re.compile(r"(?!x)x")  # TODO: Callout-Stichwort-Regex eintragen
STICHWORT_SDS_RE = re.compile(r"(?!x)x")      # TODO: SDS-Stichwort-Regex eintragen


def _extract(content: str, kind: str) -> str | None:
    pattern = STICHWORT_SDS_RE if kind == "sds" else STICHWORT_CALLOUT_RE
    m = pattern.search(content)
    return m.group(1) if m else None


EXTRACTOR = Extractor(
    name="stichwort",
    func=_extract,
    description="Stichwort (Callout und SDS, eigene Regex je kind)",
    order=50,
)
