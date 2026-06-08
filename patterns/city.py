"""Ortsname — aus strukturierter Adresse (SDS EO-Block oder Callout ||PLZ Ort||)."""

from extractor import Extractor
from patterns._address import parse_address


def _extract(content: str, kind: str) -> str | None:
    return parse_address(content, kind)["ort"]


EXTRACTOR = Extractor(
    name="city",
    func=_extract,
    description="Ort — SDS aus EO:-Block, Callout aus ||PLZ Ort||-Block (links vom ' - ')",
    order=30,
)
