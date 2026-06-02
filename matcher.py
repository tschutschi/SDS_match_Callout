"""Vergleich von Records, gewichteter Score, Union-Find-Gruppierung zu Fällen."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta

from extractor import Extractor


@dataclass
class Record:
    timestamp: datetime | None
    kind: str  # "callout" | "sds"
    content: str
    fields: dict[str, str | None] = field(default_factory=dict)


@dataclass
class Case:
    records: list[Record]
    score_max: float = 0.0
    score_avg: float = 0.0


def pair_score(
    a: Record,
    b: Record,
    weights: dict[str, float],
    extractors: list[Extractor],
) -> float:
    """Summe der Gewichte ueber Felder, in denen beide Records denselben
    normalisierten Wert haben."""
    by_name = {e.name: e for e in extractors}
    total = 0.0
    for name, weight in weights.items():
        e = by_name.get(name)
        if e is None:
            continue
        va, vb = a.fields.get(name), b.fields.get(name)
        if not va or not vb:
            continue
        if e.normalize(va) == e.normalize(vb):
            total += weight
    return total


class _UnionFind:
    def __init__(self, n: int) -> None:
        self.parent = list(range(n))

    def find(self, x: int) -> int:
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[ra] = rb


def find_cases(
    records: list[Record],
    extractors: list[Extractor],
    weights: dict[str, float],
    time_window_minutes: int,
    score_threshold: float,
) -> list[Case]:
    """Gruppiert Records zu Fällen. Singletons werden als 1-Record-Cases zurueckgegeben."""
    n = len(records)
    if n == 0:
        return []

    # Sortier-Reihenfolge: Records ohne Zeit ans Ende
    order = sorted(
        range(n),
        key=lambda i: (records[i].timestamp is None, records[i].timestamp or datetime.max),
    )
    window = timedelta(minutes=time_window_minutes)

    uf = _UnionFind(n)
    pair_scores: dict[tuple[int, int], float] = {}

    for pos, i in enumerate(order):
        ri = records[i]
        for j in order[pos + 1 :]:
            rj = records[j]
            # Zeitfenster nur anwenden, wenn beide eine Zeit haben
            if ri.timestamp is not None and rj.timestamp is not None:
                if rj.timestamp - ri.timestamp > window:
                    break  # weiter weg → wegen Sortierung auch alle danach
            s = pair_score(ri, rj, weights, extractors)
            if s >= score_threshold:
                uf.union(i, j)
                pair_scores[(min(i, j), max(i, j))] = s

    # Cluster sammeln
    clusters: dict[int, list[int]] = {}
    for idx in range(n):
        clusters.setdefault(uf.find(idx), []).append(idx)

    cases: list[Case] = []
    for members in clusters.values():
        scores = [
            pair_scores[(a, b)]
            for a in members
            for b in members
            if a < b and (a, b) in pair_scores
        ]
        cases.append(
            Case(
                records=[records[i] for i in members],
                score_max=max(scores) if scores else 0.0,
                score_avg=(sum(scores) / len(scores)) if scores else 0.0,
            )
        )

    def _sort_key(c: Case):
        ts = [r.timestamp for r in c.records if r.timestamp is not None]
        return (not ts, min(ts) if ts else datetime.max)

    cases.sort(key=_sort_key)
    return cases
