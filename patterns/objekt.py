"""Objekt — z.B. Kirche, Schule. Optional in SDS und Callout.

SDS:     ||OBJ: <Objekt>||
Callout: <Objekt>-Block direkt nach dem Straßen-Block, z.B. ...||Kirchplatz 2  ||Kirche||...

Die eigentliche Logik liegt in patterns/_address.py.
"""

from extractor import Extractor
from patterns._address import parse_address


def _extract(content: str, kind: str) -> str | None:
    return parse_address(content, kind)["objekt"]


EXTRACTOR = Extractor(
    name="objekt",
    func=_extract,
    description="Objekt — SDS aus ||OBJ:||-Block, Callout aus Block direkt nach der Straße",
    order=35,
)
