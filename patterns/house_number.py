"""Hausnummer — eigenes Suchmuster, damit '22a' vs '22' beim Match unterscheidbar bleibt.

Format: Zahl + optional 1-2 Kleinbuchstaben + optional Bruch ('/' oder '-').
Beispiele: 5 / 22 / 22a / 12ab / 22/24 / 5-7
"""

from extractor import Extractor
from patterns._address import parse_address


def _extract(content: str, kind: str) -> str | None:
    return parse_address(content, kind)["hausnummer"]


EXTRACTOR = Extractor(
    name="house_number",
    func=_extract,
    description="Hausnummer (Zahl + optional Buchstaben/Bruch)",
    order=25,
)
