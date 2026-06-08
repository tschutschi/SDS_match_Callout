"""Gemeinsamer Adress-Parser für SDS und Callout.

SDS-Format:
    ||EO: 82515 Wolfratshausen; Margeritenstraße; 22a||

Callout-Format:
    ||82515 Wolfratshausen - Wolfratshausen||Margeritenstraße 22a  og 2||

Die einzelnen Pattern-Dateien (plz, city, street, house_number) holen sich
hier ihren Wert ab — parse_address ist gecacht, also läuft das Parsen pro
Zeile nur einmal.
"""

from __future__ import annotations

import re
from functools import lru_cache

# Hausnummer DE: Zahl + optional 1-2 Kleinbuchstaben + optional Bruch (/ oder -)
HAUSNR_PATTERN = (
    r"\d+\s?[a-zäöüß]{0,2}"
    r"(?:\s?[/\-]\s?\d+\s?[a-zäöüß]{0,2})?"
)

_SDS_ADDRESS_RE = re.compile(
    r"\|\|\s*EO:\s*"
    r"(?P<plz>\d{5})\s+"
    r"(?P<ort>[^;|]+?)\s*;\s*"
    r"(?P<strasse>[^;|]+?)\s*;\s*"
    r"(?P<hausnummer>[^|]+?)\s*\|\|",
    re.IGNORECASE,
)

# Callout-Vollblock: ||PLZ Ort - Gemeinde||Straße Hausnummer  Zusatz||
# Strassen-Block muss DIREKT nach dem PLZ-Ort-Block kommen — sonst koennten
# Einsatzmittel-Blöcke wie '||FL Sleh 11/1||' faelschlich als Strasse gelten.
_CO_FULL_ADDRESS_RE = re.compile(
    r"\|\|(?P<plz>\d{5})\s+(?P<ort_raw>[^|]+?)\|\|"
    r"(?P<strasse>[A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß.\-]*"
    r"(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*?)"
    r"\s+(?P<hausnummer>" + HAUSNR_PATTERN + r")\b"
)

# Fallback: nur ||PLZ Ort|| (falls kein Strasse-Block folgt)
_CO_PLZ_ORT_ONLY_RE = re.compile(r"\|\|(?P<plz>\d{5})\s+(?P<ort_raw>[^|]+?)\|\|")


_EMPTY = {"plz": None, "ort": None, "strasse": None, "hausnummer": None}


def _co_ort_from_raw(ort_raw: str) -> str:
    """'Ortsname - Gemeinde' → 'Ortsname'."""
    return ort_raw.strip().split(" - ", 1)[0].strip()


@lru_cache(maxsize=2048)
def parse_address(content: str, kind: str) -> dict[str, str | None]:
    """Liefert {'plz', 'ort', 'strasse', 'hausnummer'} oder Felder mit None."""
    if kind == "sds":
        m = _SDS_ADDRESS_RE.search(content)
        if not m:
            return dict(_EMPTY)
        return {k: v.strip() for k, v in m.groupdict().items()}

    if kind == "callout":
        out = dict(_EMPTY)
        m = _CO_FULL_ADDRESS_RE.search(content)
        if m:
            out["plz"] = m.group("plz")
            out["ort"] = _co_ort_from_raw(m.group("ort_raw"))
            out["strasse"] = m.group("strasse").strip()
            out["hausnummer"] = re.sub(r"\s+", "", m.group("hausnummer"))
            return out
        m1 = _CO_PLZ_ORT_ONLY_RE.search(content)
        if m1:
            out["plz"] = m1.group("plz")
            out["ort"] = _co_ort_from_raw(m1.group("ort_raw"))
        return out

    return dict(_EMPTY)
