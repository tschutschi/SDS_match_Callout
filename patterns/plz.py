"""Postleitzahl — aus strukturierter Adresse."""

from extractor import Extractor
from patterns._address import parse_address


def _extract(content: str, kind: str) -> str | None:
    return parse_address(content, kind)["plz"]


EXTRACTOR = Extractor(
    name="plz",
    func=_extract,
    description="PLZ — SDS aus EO:-Block, Callout aus ||PLZ Ort||-Block",
    order=10,
)
