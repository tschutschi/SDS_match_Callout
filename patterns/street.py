"""Straße + Hausnummer direkt vor der PLZ."""

import re

from extractor import Extractor
from patterns.plz import PLZ_RE

STREET_END_RE = re.compile(
    r"([A-Za-zÄÖÜäöüß.\-]+(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*\s+\d+\s*[a-zA-Z]?)\s*$"
)


def _extract(content: str, kind: str) -> str | None:
    text = " ".join(content.split())
    plz_match = PLZ_RE.search(text)
    if not plz_match:
        return None
    before = text[: plz_match.start()].rstrip(" ,;:-")
    m = STREET_END_RE.search(before)
    return m.group(1).strip() if m else None


EXTRACTOR = Extractor(
    name="street",
    func=_extract,
    description="Straße + Hausnummer vor der PLZ",
    order=20,
)
