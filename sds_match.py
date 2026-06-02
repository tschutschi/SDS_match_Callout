#!/usr/bin/env python3
"""SDS/Callout — Excel einlesen, klassifizieren, Suchmuster anwenden,
Records zu Fällen gruppieren (gewichteter Score, Zeitfenster).

Logik & Mechanik. Suchmuster liegen in patterns/*.py, Klassifikator-
Praefix + Gewichte + Zeitfenster in config.yaml.
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

from config import load_config
from extractor import Extractor, apply_extractors, classify, load_patterns
from matcher import Case, Record, find_cases


# ---------- Excel ----------

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
    sheet,
    extractors: list[Extractor],
    callout_prefix: str,
) -> list[Record]:
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=None, dtype=str)
    df = df.iloc[:, :3]
    df.columns = ["date", "time", "content"]

    records: list[Record] = []
    for _, row in df.iterrows():
        content = "" if pd.isna(row["content"]) else str(row["content"])
        if not content.strip():
            continue
        kind = classify(content, callout_prefix)
        records.append(
            Record(
                timestamp=parse_timestamp(row["date"], row["time"]),
                kind=kind,
                content=content,
                fields=apply_extractors(content, kind, extractors),
            )
        )
    return records


# ---------- CLI Helpers ----------

def configure_extractors(
    extractors: list[Extractor],
    only: str | None,
    skip: str | None,
) -> list[Extractor]:
    names = {e.name for e in extractors}

    if only:
        wanted = {n.strip() for n in only.split(",") if n.strip()}
        unknown = wanted - names
        if unknown:
            raise SystemExit(f"Unbekannte Suchmuster: {sorted(unknown)}. Verfügbar: {sorted(names)}")
        for e in extractors:
            e.enabled = e.name in wanted

    if skip:
        unwanted = {n.strip() for n in skip.split(",") if n.strip()}
        unknown = unwanted - names
        if unknown:
            raise SystemExit(f"Unbekannte Suchmuster: {sorted(unknown)}. Verfügbar: {sorted(names)}")
        for e in extractors:
            if e.name in unwanted:
                e.enabled = False

    return [e for e in extractors if e.enabled]


def _fmt_ts(ts) -> str:
    return ts.isoformat(sep=" ", timespec="milliseconds") if ts else "-"


# ---------- Output ----------

def print_cases(
    cases: list[Case],
    active: list[Extractor],
    show_content: bool,
    only_multi: bool,
) -> None:
    shown = 0
    case_no = 0
    for c in cases:
        if only_multi and len(c.records) < 2:
            continue
        case_no += 1
        shown += 1
        n_sds = sum(1 for r in c.records if r.kind == "sds")
        n_co = sum(1 for r in c.records if r.kind == "callout")
        parts = []
        if n_co:
            parts.append(f"{n_co} Callout" + ("s" if n_co != 1 else ""))
        if n_sds:
            parts.append(f"{n_sds} SDS")
        head = ", ".join(parts) if parts else "0"
        suffix = "" if len(c.records) > 1 else ", ohne Gegenstück"
        print(
            f"\nFall #{case_no}  {len(c.records)} Record(s) ({head}{suffix})  "
            f"score_max={c.score_max:.2f}  score_avg={c.score_avg:.2f}"
        )
        for r in c.records:
            field_parts = [
                f"{e.name}={r.fields.get(e.name) or '-'!r}" for e in active
            ]
            print(f"  [{r.kind:7s}] {_fmt_ts(r.timestamp)}  " + "  ".join(field_parts))
            if show_content:
                print(f"           {r.content}")
    if only_multi and shown == 0:
        print("\n(Keine Fälle mit >= 2 Records gefunden.)")


def print_stats(records: list[Record], active: list[Extractor]) -> None:
    n_co = sum(1 for r in records if r.kind == "callout")
    n_sds = sum(1 for r in records if r.kind == "sds")
    print("\nMatch-Statistik (pro Suchmuster):")
    print(f"  {'Muster':12s}  {'Callout':>14s}  {'SDS':>14s}  {'Gesamt':>14s}")
    for e in active:
        co_hit = sum(1 for r in records if r.kind == "callout" and r.fields.get(e.name))
        sds_hit = sum(1 for r in records if r.kind == "sds" and r.fields.get(e.name))
        total_hit = co_hit + sds_hit
        co_str = f"{co_hit}/{n_co}" if n_co else "-"
        sds_str = f"{sds_hit}/{n_sds}" if n_sds else "-"
        tot_str = f"{total_hit}/{len(records)}"
        print(f"  {e.name:12s}  {co_str:>14s}  {sds_str:>14s}  {tot_str:>14s}")


# ---------- Main ----------

def main() -> None:
    p = argparse.ArgumentParser(description="SDS/Callout — Fälle gruppieren")
    p.add_argument("xlsx", type=Path, nargs="?", help="Pfad zur Excel-Datei")
    p.add_argument("--sheet", default=0, help="Sheetname oder -index (Default 0)")
    p.add_argument("--config", type=Path, default=Path("config.yaml"))
    p.add_argument("--only", help="Nur diese Suchmuster aktivieren (kommagetrennt)")
    p.add_argument("--skip", help="Diese Suchmuster deaktivieren (kommagetrennt)")
    p.add_argument("--list", action="store_true", help="Verfügbare Suchmuster auflisten und beenden")
    p.add_argument("--show-content", action="store_true", help="Original-Text mit ausgeben")
    p.add_argument("--threshold", type=float, help="Score-Schwellwert (überschreibt Config)")
    p.add_argument("--time-window", type=int, help="Zeitfenster Minuten (überschreibt Config)")
    p.add_argument("--only-multi", action="store_true", help="Nur Fälle mit >=2 Records anzeigen")
    args = p.parse_args()

    extractors = load_patterns()

    if args.list:
        print("Verfügbare Suchmuster:")
        for e in extractors:
            print(f"  {e.name:12s}  {e.description}")
        return

    if not args.xlsx:
        p.error("Pfad zur Excel-Datei fehlt (oder --list verwenden)")

    cfg = load_config(args.config if args.config.exists() else None)
    threshold = args.threshold if args.threshold is not None else cfg.score_threshold
    window = args.time_window if args.time_window is not None else cfg.time_window_minutes

    active = configure_extractors(extractors, args.only, args.skip)
    if not active:
        raise SystemExit("Keine Suchmuster aktiv — nichts zu tun.")

    known = {e.name for e in extractors}
    unknown_w = set(cfg.weights) - known
    if unknown_w:
        print(f"Warnung: Gewichte für unbekannte Suchmuster: {sorted(unknown_w)}")
    inactive_w = (set(cfg.weights) & known) - {e.name for e in active}
    if inactive_w:
        print(f"Hinweis: Gewichte für deaktivierte Suchmuster werden ignoriert: {sorted(inactive_w)}")

    effective_weights = {n: w for n, w in cfg.weights.items() if n in {e.name for e in active}}

    records = load_records(args.xlsx, args.sheet, active, cfg.callout_prefix)
    n_co = sum(1 for r in records if r.kind == "callout")
    n_sds = sum(1 for r in records if r.kind == "sds")
    print(f"Gelesen: {len(records)} Datensätze ({n_co} Callouts, {n_sds} SDS)")
    print(f"Aktive Suchmuster: {', '.join(e.name for e in active)}")
    print(f"Zeitfenster: {window} min   Schwellwert: {threshold:.2f}")

    cases = find_cases(
        records,
        extractors=active,
        weights=effective_weights,
        time_window_minutes=window,
        score_threshold=threshold,
    )

    multi = sum(1 for c in cases if len(c.records) >= 2)
    single = sum(1 for c in cases if len(c.records) == 1)
    print(f"Fälle: {len(cases)} ({multi} mit mehreren Records, {single} Singletons)")

    print_cases(cases, active, args.show_content, args.only_multi)
    print_stats(records, active)


if __name__ == "__main__":
    main()
