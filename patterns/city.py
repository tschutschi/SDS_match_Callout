"""Ortsname nach der PLZ — getrennte Regex für Callout und SDS."""

import re

from extractor import Extractor
from patterns.plz import PLZ_RE

# Start identisch; spaeter unabhaengig tunbar
CITY_CALLOUT_RE = re.compile(
    r"([A-Za-zÄÖÜäöüß.\-]+(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*)"
)
CITY_SDS_RE = re.compile(
    r"([A-Za-zÄÖÜäöüß.\-]+(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*)"
)


def _extract(content: str, kind: str) -> str | None:
    text = " ".join(content.split())
    plz_match = PLZ_RE.search(text)
    if not plz_match:
        return None
    after = text[plz_match.end() :].lstrip(" ,;:-")
    pattern = CITY_SDS_RE if kind == "sds" else CITY_CALLOUT_RE
    m = pattern.match(after)
    return m.group(1).strip() if m else None


EXTRACTOR = Extractor(
    name="city",
    func=_extract,
    description="Ortsname nach der PLZ (kind-spezifisch)",
    order=30,
)
