"""Ortsname — aus strukturierter Adresse (SDS EO-Block oder Callout ||PLZ Ort||)."""

from extractor import Extractor

# Callout: PLZ steht direkt vor dem Ort (z.B. "82418 Murnau").
# Die PLZ-Form ist hier nur Anker — die Erkennung der PLZ als Suchmuster
# liegt weiterhin in patterns/plz.py.
CITY_CALLOUT_RE = re.compile(
    r"\b8(?:2|3)\d{3}\s+([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß.\-]*(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*)"
)

# SDS: kein PLZ-Kontext. Hier Deine SDS-spezifische Regex eintragen.
# Solange Platzhalter (matcht nie), gibt es fuer SDS-Zeilen keine Stadt.
CITY_SDS_RE = re.compile(r"\|\|OT: ([A-Za-zÄÖÜäöüß.\-]+(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*)")  # TODO: SDS-Ort-Regex eintragen


def _extract(content: str, kind: str) -> str | None:
    return parse_address(content, kind)["ort"]


EXTRACTOR = Extractor(
    name="city",
    func=_extract,
    description="Ort — SDS aus EO:-Block, Callout aus ||PLZ Ort||-Block (links vom ' - ')",
    order=30,
)
