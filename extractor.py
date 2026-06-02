"""Extractor-Framework: Suchmuster aus patterns/ laden + Klassifikation."""

from __future__ import annotations

import importlib
import pkgutil
import re
from dataclasses import dataclass, field
from typing import Callable

ExtractFn = Callable[[str, str], str | None]
NormalizeFn = Callable[[str], str]


_STR_RE = re.compile(r"\bstr\.?\b", flags=re.IGNORECASE)
_STRASSE_LATIN_RE = re.compile(r"\bstrasse\b", flags=re.IGNORECASE)
_WS_RE = re.compile(r"\s+")


def default_normalize(value: str) -> str:
    """strip + lower + Whitespace zusammenfassen + 'str.' / 'strasse' -> 'straße'."""
    v = value.strip().lower()
    v = _STR_RE.sub("straße", v)
    v = _STRASSE_LATIN_RE.sub("straße", v)
    v = _WS_RE.sub(" ", v)
    return v


@dataclass
class Extractor:
    name: str
    func: ExtractFn
    description: str = ""
    normalize: NormalizeFn = field(default=default_normalize)
    order: int = 100
    enabled: bool = True


def load_patterns() -> list[Extractor]:
    """Sammelt alle EXTRACTOR-Variablen aus patterns/*.py.
    Sortierung: nach .order, dann .name."""
    import patterns as patterns_pkg

    out: list[Extractor] = []
    for mod_info in pkgutil.iter_modules(patterns_pkg.__path__):
        mod = importlib.import_module(f"patterns.{mod_info.name}")
        extractor = getattr(mod, "EXTRACTOR", None)
        if isinstance(extractor, Extractor):
            out.append(extractor)
    out.sort(key=lambda e: (e.order, e.name))
    return out


def apply_extractors(
    content: str,
    kind: str,
    extractors: list[Extractor],
) -> dict[str, str | None]:
    return {e.name: e.func(content, kind) for e in extractors if e.enabled}


def classify(content: str, callout_prefix: str) -> str:
    return "callout" if content.lstrip().startswith(callout_prefix) else "sds"
