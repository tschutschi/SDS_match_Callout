"""Postleitzahl — 5-stellig, beginnend mit 82 oder 83."""

import re

from extractor import Extractor

PLZ_RE = re.compile(r"\b(8(2|3)\d{3})\b")


def _extract(content: str, kind: str) -> str | None:
    m = PLZ_RE.search(content)
    return m.group(1) if m else None


EXTRACTOR = Extractor(
    name="plz",
    func=_extract,
    description="Postleitzahl (82xxx / 83xxx)",
    order=10,
)
