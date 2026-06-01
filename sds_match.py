#!/usr/bin/env python3
"""Stage 1: Excel einlesen, Zeilen als SDS/Callout klassifizieren,
und konfigurierbare Suchmuster auf den Inhalt anwenden.

Neue Suchmuster:
  1. Funktion definieren: extract_<name>(content: str, kind: str) -> str | None
  2. In EXTRACTORS-Liste am Ende der Datei eintragen
  3. Fertig — taucht automatisch in --list, --only, --skip und Output auf
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable

import pandas as pd

CALLOUT_PREFIX = "IncomingCallout:"

# ----- gemeinsame Regex (von mehreren Extraktoren genutzt) -----
PLZ_RE = re.compile(r"\b(8(2|3)\d{3})\b")
STREET_END_RE = re.compile(
    r"([A-Za-zÄÖÜäöüß.\-]+(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*\s+\d+\s*[a-zA-Z]?)\s*$"
)
CITY_START_RE = re.compile(
    r"([A-Za-zÄÖÜäöüß.\-]+(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*)"
)
SCHLAGWORT_CALLOUT_RE = re.compile(r"\|(#[TBIR]\d{4}[a-zA-ZÖöÄäß#\-\ ]*)\|+")
SCHLAGWORT_SDS_RE = re.compile(r"\|+SW: (#[BTIR]\d{4}[a-zA-ZÖöÄäß#\-\ ]*)\|+")


# ============================================================
#  Extractor-Framework
# ============================================================

ExtractFn = Callable[[str, str], str | None]


@dataclass
class Extractor:
    name: str
    func: ExtractFn
    description: str = ""
    enabled: bool = True


@dataclass
class Record:
    timestamp: datetime | None
    kind: str  # "callout" | "sds"
    content: str
    fields: dict[str, str | None] = field(default_factory=dict)


def classify(content: str) -> str:
    return "callout" if content.lstrip().startswith(CALLOUT_PREFIX) else "sds"


# ============================================================
#  Einzelne Suchmuster (Extraktoren)
# ============================================================

def extract_plz(content: str, kind: str) -> str | None:
    m = PLZ_RE.search(content)
    return m.group(1) if m else None


def extract_street(content: str, kind: str) -> str | None:
    text = " ".join(content.split())
    plz_match = PLZ_RE.search(text)
    if not plz_match:
        return None
    before = text[: plz_match.start()].rstrip(" ,;:-")
    m = STREET_END_RE.search(before)
    return m.group(1).strip() if m else None


def extract_city(content: str, kind: str) -> str | None:
    text = " ".join(content.split())
    plz_match = PLZ_RE.search(text)
    if not plz_match:
        return None
    after = text[plz_match.end() :].lstrip(" ,;:-")
    m = CITY_START_RE.match(after)
    return m.group(1).strip() if m else None


def extract_schlagwort(content: str, kind: str) -> str | None:
    pattern = SCHLAGWORT_SDS_RE if kind == "sds" else SCHLAGWORT_CALLOUT_RE
    m = pattern.search(content)
    return m.group(1) if m else None


# ============================================================
#  Registry — hier neue Suchmuster eintragen
# ============================================================

EXTRACTORS: list[Extractor] = [
    Extractor("plz",        extract_plz,        "Postleitzahl (82xxx)"),
    Extractor("street",     extract_street,     "Straße + Hausnummer vor der PLZ"),
    Extractor("city",       extract_city,       "Ortsname nach der PLZ"),
    Extractor("schlagwort", extract_schlagwort, "Schlagwort zwischen Pipes (B#/T#/I#)"),
]


def apply_extractors(content: str, kind: str, extractors: list[Extractor]) -> dict[str, str | None]:
    return {e.name: e.func(content, kind) for e in extractors if e.enabled}


# ============================================================
#  Excel laden
# ============================================================

def parse_timestamp(date_val, time_val) -> datetime | None:
    if pd.isna(date_val) or pd.isna(time_val):
        return None
    date_str = str(date_val).strip().lstrip("'")
    time_str = str(time_val).strip().lstrip("'").replace(",", ".")
    for fmt in ("%d.%m.%Y %H:%M:%S.%f", "%d.%m.%Y %H:%M:%S"):
        try:
            return datetime.strptime(f"{date_str} {time_str}", fmt)
        except ValueError:
            continue
    return None


def load_records(
    xlsx_path: Path,
    sheet: str | int,
    extractors: list[Extractor],
) -> list[Record]:
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=None, dtype=str)
    df = df.iloc[:, :3]
    df.columns = ["date", "time", "content"]

    records: list[Record] = []
    for _, row in df.iterrows():
        content = "" if pd.isna(row["content"]) else str(row["content"])
        if not content.strip():
            continue
        kind = classify(content)
        records.append(
            Record(
                timestamp=parse_timestamp(row["date"], row["time"]),
                kind=kind,
                content=content,
                fields=apply_extractors(content, kind, extractors),
            )
        )
    return records


# ============================================================
#  CLI
# ============================================================

def configure_extractors(only: str | None, skip: str | None) -> list[Extractor]:
    names = {e.name for e in EXTRACTORS}

    if only:
        wanted = {n.strip() for n in only.split(",") if n.strip()}
        unknown = wanted - names
        if unknown:
            raise SystemExit(f"Unbekannte Suchmuster: {sorted(unknown)}. Verfügbar: {sorted(names)}")
        for e in EXTRACTORS:
            e.enabled = e.name in wanted

    if skip:
        unwanted = {n.strip() for n in skip.split(",") if n.strip()}
        unknown = unwanted - names
        if unknown:
            raise SystemExit(f"Unbekannte Suchmuster: {sorted(unknown)}. Verfügbar: {sorted(names)}")
        for e in EXTRACTORS:
            if e.name in unwanted:
                e.enabled = False

    return [e for e in EXTRACTORS if e.enabled]


def main() -> None:
    p = argparse.ArgumentParser(description="SDS/Callout Stage 1 – einlesen & parsen")
    p.add_argument("xlsx", type=Path, nargs="?", help="Pfad zur Excel-Datei")
    p.add_argument("--sheet", default=0, help="Sheetname oder -index (Default 0)")
    p.add_argument("--limit", type=int, default=20, help="Wie viele Zeilen ausgeben (-1 = alle)")
    p.add_argument("--only", help="Nur diese Suchmuster aktivieren (kommagetrennt)")
    p.add_argument("--skip", help="Diese Suchmuster deaktivieren (kommagetrennt)")
    p.add_argument("--list", action="store_true", help="Verfügbare Suchmuster auflisten und beenden")
    p.add_argument("--show-content", action="store_true", help="Original-Text mit ausgeben")
    args = p.parse_args()

    if args.list:
        print("Verfügbare Suchmuster:")
        for e in EXTRACTORS:
            print(f"  {e.name:12s}  {e.description}")
        return

    if not args.xlsx:
        p.error("Pfad zur Excel-Datei fehlt (oder --list verwenden)")

    active = configure_extractors(args.only, args.skip)
    if not active:
        raise SystemExit("Keine Suchmuster aktiv — nichts zu tun.")

    records = load_records(args.xlsx, args.sheet, active)
    n_co = sum(1 for r in records if r.kind == "callout")
    n_sds = sum(1 for r in records if r.kind == "sds")
    print(f"Gelesen: {len(records)} Datensätze ({n_co} Callouts, {n_sds} SDS)")
    print(f"Aktive Suchmuster: {', '.join(e.name for e in active)}\n")

    limit = len(records) if args.limit < 0 else args.limit
    shown = records[:limit]
    for r in shown:
        ts = r.timestamp.isoformat(sep=" ", timespec="milliseconds") if r.timestamp else "-"
        parts = [f"{name}={r.fields.get(name) or '-'!r}" for name in (e.name for e in active)]
        line = f"[{r.kind:7s}] {ts}  " + "  ".join(parts)
        if args.show_content:
            line += f"\n           {r.content}"
        print(line)
    if len(records) > len(shown):
        print(f"... ({len(records) - len(shown)} weitere Zeilen — mit --limit -1 alle anzeigen)")

    # ----- Statistik je Suchmuster, getrennt nach SDS/Callout -----
    print("\nMatch-Statistik (pro Suchmuster):")
    print(f"  {'Muster':12s}  {'Callout':>14s}  {'SDS':>14s}  {'Gesamt':>14s}")
    for e in active:
        co_hit = sum(1 for r in records if r.kind == "callout" and r.fields.get(e.name))
        sds_hit = sum(1 for r in records if r.kind == "sds"     and r.fields.get(e.name))
        total_hit = co_hit + sds_hit
        co_str  = f"{co_hit}/{n_co}"  if n_co  else "-"
        sds_str = f"{sds_hit}/{n_sds}" if n_sds else "-"
        tot_str = f"{total_hit}/{len(records)}"
        print(f"  {e.name:12s}  {co_str:>14s}  {sds_str:>14s}  {tot_str:>14s}")

    # ----- Records ohne jeden Match (Hinweis fuer Regex-Tuning) -----
    no_match = [r for r in records if not any(r.fields.values())]
    if no_match:
        print(f"\n{len(no_match)} Datensätze ohne irgendeinen Treffer:")
        for r in no_match[:10]:
            ts = r.timestamp.isoformat(sep=" ", timespec="milliseconds") if r.timestamp else "-"
            print(f"  [{r.kind:7s}] {ts}  {r.content[:120]}")
        if len(no_match) > 10:
            print(f"  ... ({len(no_match) - 10} weitere)")


if __name__ == "__main__":
    main()
