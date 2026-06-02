"""Config laden — YAML mit Defaults."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml

DEFAULTS = {
    "classification": {
        "callout_prefix": "IncomingCallout:",
    },
    "matching": {
        "time_window_minutes": 120,
        "score_threshold": 0.5,
    },
    "weights": {},
}


@dataclass
class Config:
    callout_prefix: str
    time_window_minutes: int
    score_threshold: float
    weights: dict[str, float]


def _deep_merge(base: dict, overlay: dict) -> dict:
    out = dict(base)
    for k, v in overlay.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(path: Path | None) -> Config:
    raw = DEFAULTS
    if path and path.exists():
        with path.open() as fh:
            user = yaml.safe_load(fh) or {}
        raw = _deep_merge(DEFAULTS, user)
    return Config(
        callout_prefix=str(raw["classification"]["callout_prefix"]),
        time_window_minutes=int(raw["matching"]["time_window_minutes"]),
        score_threshold=float(raw["matching"]["score_threshold"]),
        weights={str(k): float(v) for k, v in (raw.get("weights") or {}).items()},
    )
