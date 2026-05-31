#!/usr/bin/env python3
"""Stage 1: Excel einlesen, Zeilen als SDS/Callout klassifizieren,
PLZ/Straße/Ort aus dem Freitext extrahieren."""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

CALLOUT_PREFIX = "IncomingCallout:"

# PLZen in unserem Gebiet beginnen mit 82
PLZ_RE = re.compile(r"\b(82\d{3})\b")

# Straßenname + Hausnummer (optional Buchstaben-Zusatz wie 12a)
STREET_END_RE = re.compile(
    r"([A-Za-zÄÖÜäöüß.\-]+(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*\s+\d+\s*[a-zA-Z]?)\s*$"
)
# Ortsname direkt nach der PLZ
CITY_START_RE = re.compile(
    r"([A-Za-zÄÖÜäöüß.\-]+(?:[ \-][A-Za-zÄÖÜäöüß.\-]+)*)"
)


@dataclass
class Address:
    raw: str
    plz: str | None
    street: str | None
    city: str | None


@dataclass
class Record:
    timestamp: datetime | None
    kind: str  # "callout" | "sds"
    content: str
    address: Address


def classify(content: str) -> str:
    return "callout" if content.lstrip().startswith(CALLOUT_PREFIX) else "sds"


def parse_address(content: str) -> Address:
    text = " ".join(content.split())
    plz_match = PLZ_RE.search(text)
    if not plz_match:
        return Address(raw=content, plz=None, street=None, city=None)

    plz = plz_match.group(1)
    before = text[: plz_match.start()].rstrip(" ,;:-")
    after = text[plz_match.end() :].lstrip(" ,;:-")

    street_match = STREET_END_RE.search(before)
    street = street_match.group(1).strip() if street_match else None

    city_match = CITY_START_RE.match(after)
    city = city_match.group(1).strip() if city_match else None

    return Address(raw=content, plz=plz, street=street, city=city)


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


def load_records(xlsx_path: Path, sheet: str | int = 0) -> list[Record]:
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=None, dtype=str)
    df = df.iloc[:, :3]
    df.columns = ["date", "time", "content"]

    records: list[Record] = []
    for _, row in df.iterrows():
        content = "" if pd.isna(row["content"]) else str(row["content"])
        if not content.strip():
            continue
        records.append(
            Record(
                timestamp=parse_timestamp(row["date"], row["time"]),
                kind=classify(content),
                content=content,
                address=parse_address(content),
            )
        )
    return records


def main() -> None:
    p = argparse.ArgumentParser(description="SDS/Callout Stage 1 – einlesen & parsen")
    p.add_argument("xlsx", type=Path, help="Pfad zur Excel-Datei")
    p.add_argument("--sheet", default=0, help="Sheetname oder -index (Default 0)")
    p.add_argument("--limit", type=int, default=20, help="Wie viele Zeilen ausgeben")
    args = p.parse_args()

    records = load_records(args.xlsx, args.sheet)
    n_co = sum(1 for r in records if r.kind == "callout")
    n_sds = sum(1 for r in records if r.kind == "sds")
    print(f"Gelesen: {len(records)} Datensätze ({n_co} Callouts, {n_sds} SDS)\n")

    for r in records[: args.limit]:
        ts = r.timestamp.isoformat(sep=" ", timespec="milliseconds") if r.timestamp else "-"
        print(
            f"[{r.kind:7s}] {ts}  "
            f"PLZ={r.address.plz or '-':5s}  "
            f"Str={r.address.street or '-'!r:40s}  "
            f"Ort={r.address.city or '-'!r}"
        )


if __name__ == "__main__":
    main()
