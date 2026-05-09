#!/usr/bin/env python3
"""Materialize the calamine NYC 311 CSV benchmark source as XLSX.

The upstream calamine README benchmarks an XLSX workbook, while the published
download is a compressed CSV. This script makes the CSV-to-XLSX step explicit
and deterministic enough for local benchmark artifact pinning.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import time
from datetime import datetime, timezone
from pathlib import Path

import xlsxwriter

SHEET_NAME = "NYC_311_SR_2010-2020-sample-1M"
EXPECTED_ROWS = 1_000_001
EXPECTED_COLS = 41
EXPECTED_NONEMPTY = 28_056_975
EXPECTED_CSV_SHA256 = "18f0dd774a6c4b79da3dbf3aa0cd878d374dab132226af2c629d9eef9595061b"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source_csv", type=Path)
    parser.add_argument("output_xlsx", type=Path)
    parser.add_argument("--progress-every", type=int, default=100_000)
    parser.add_argument("--json", action="store_true")
    return parser.parse_args()


def emit(payload: dict[str, object], as_json: bool) -> None:
    if as_json:
        print(json.dumps(payload, sort_keys=True), flush=True)
    else:
        print(payload, flush=True)


def main() -> None:
    args = parse_args()
    source_csv = args.source_csv
    output_xlsx = args.output_xlsx
    csv_sha256 = sha256_file(source_csv)
    if csv_sha256 != EXPECTED_CSV_SHA256:
        raise SystemExit(f"unexpected CSV sha256: {csv_sha256}")

    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    start = time.perf_counter()
    workbook = xlsxwriter.Workbook(
        str(output_xlsx),
        {"constant_memory": True, "strings_to_numbers": True},
    )
    workbook.set_properties(
        {
            "title": "NYC 311 SR 2010-2020 sample 1M",
            "subject": "calamine upstream real-workbook benchmark fixture",
            "author": "Ascend benchmark materializer",
            "comments": "Generated from the pinned upstream CSV source.",
            "created": datetime(2021, 11, 5, tzinfo=timezone.utc),
        }
    )
    worksheet = workbook.add_worksheet(SHEET_NAME)

    rows = 0
    cols = 0
    nonempty = 0
    with source_csv.open("r", newline="", encoding="utf-8") as handle:
        reader = csv.reader(handle)
        for row_index, row in enumerate(reader):
            rows = row_index + 1
            cols = max(cols, len(row))
            for col_index, value in enumerate(row):
                if value == "":
                    continue
                worksheet.write(row_index, col_index, value)
                nonempty += 1
            if args.progress_every > 0 and rows % args.progress_every == 0:
                emit(
                    {
                        "rows": rows,
                        "nonempty": nonempty,
                        "elapsedSec": round(time.perf_counter() - start, 2),
                    },
                    args.json,
                )

    workbook.close()
    if rows != EXPECTED_ROWS or cols != EXPECTED_COLS or nonempty != EXPECTED_NONEMPTY:
        raise SystemExit(
            "unexpected CSV shape: "
            f"rows={rows} cols={cols} nonempty={nonempty}"
        )
    emit(
        {
            "output": str(output_xlsx),
            "rows": rows,
            "cols": cols,
            "nonempty": nonempty,
            "bytes": output_xlsx.stat().st_size,
            "sha256": sha256_file(output_xlsx),
            "elapsedSec": round(time.perf_counter() - start, 2),
        },
        args.json,
    )


if __name__ == "__main__":
    main()
