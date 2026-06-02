"""Schlagwort zwischen Pipes — Callout: |#[TBIR]dddd...|, SDS: |+SW: #[BTIR]dddd...|+."""

import re

from extractor import Extractor

SCHLAGWORT_CALLOUT_RE = re.compile(r"\|(#[TBIR]\d{4}[a-zA-ZÖöÄäß#\-\ ]*)\|+")
SCHLAGWORT_SDS_RE = re.compile(r"\|+SW: (#[BTIR]\d{4}[a-zA-ZÖöÄäß#\-\ ]*)\|+")


def _extract(content: str, kind: str) -> str | None:
    pattern = SCHLAGWORT_SDS_RE if kind == "sds" else SCHLAGWORT_CALLOUT_RE
    m = pattern.search(content)
    return m.group(1) if m else None


EXTRACTOR = Extractor(
    name="schlagwort",
    func=_extract,
    description="Schlagwort zwischen Pipes (B#/T#/I#/R# + 4 Ziffern)",
    order=40,
)
