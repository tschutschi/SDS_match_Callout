"""Ortsname — aus strukturierter Adresse.

SDS:     ||OT: <Ort>||  (Vorrang) bzw. Ort aus dem ||EO: ...||-Block
Callout: ||PLZ Ort - Gemeinde||  (Ort links vom ' - ')

Die eigentliche Logik liegt in patterns/_address.py.
"""

from extractor import Extractor
from patterns._address import parse_address


def _extract(content: str, kind: str) -> str | None:
    return parse_address(content, kind)["ort"]


EXTRACTOR = Extractor(
    name="city",
    func=_extract,
    description="Ort — SDS aus ||OT:|| bzw. ||EO:||-Block, Callout aus ||PLZ Ort||-Block",
    order=30,
)
