"""Straße — aus strukturierter Adresse, OHNE Hausnummer (die ist ein eigenes Pattern)."""

from extractor import Extractor
from patterns._address import parse_address


def _extract(content: str, kind: str) -> str | None:
    return parse_address(content, kind)["strasse"]


EXTRACTOR = Extractor(
    name="street",
    func=_extract,
    description="Straße ohne Hausnummer — SDS aus EO:-Block, Callout aus ||Straße Hausnr||",
    order=20,
)
