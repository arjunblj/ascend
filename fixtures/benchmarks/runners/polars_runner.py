#!/usr/bin/env python3
"""polars runner for competitive real-workbook read benchmarks."""

from __future__ import annotations

import argparse
import hashlib
import json
import time
from pathlib import Path
from typing import Any

import polars as pl
from memory_metrics import memory_baseline, sample_with_memory


def column_name(index: int) -> str:
    name = ""
    current = index
    while current > 0:
        current, rem = divmod(current - 1, 26)
        name = chr(65 + rem) + name
    return name or "A"


def hash_lines(lines: list[str]) -> str:
    digest = hashlib.sha256()
    for line in sorted(lines):
        digest.update(f"{len(line)}:".encode("utf-8"))
        digest.update(line.encode("utf-8"))
        digest.update(b"\n")
    return digest.hexdigest()


def used_range(sheet_name: str, width: int, height: int) -> str:
    if width <= 0 or height <= 0:
        return f"{sheet_name}!empty"
    return f"{sheet_name}!A1:{column_name(width)}{height}"


def canonical_number(value: Any) -> str:
    number = float(value)
    if number == 0:
        return "0"
    if number.is_integer():
        return str(int(number))
    return format(number, ".15g")


def scalar_payload(value: Any) -> str:
    if value is None:
        return "empty"
    if isinstance(value, bool):
        return f"b:{'true' if value else 'false'}"
    if isinstance(value, (int, float)):
        return f"n:{canonical_number(value)}"
    return f"s:{value}"


def read_materialized(path: Path, engine: str) -> dict[str, pl.DataFrame]:
    result = pl.read_excel(path, has_header=False, sheet_id=0, engine=engine)
    if isinstance(result, dict):
        return result
    return {"Sheet1": result}


def read_assertions(sheets: dict[str, pl.DataFrame]) -> dict[str, str | int | bool | None]:
    cell_count = 0
    used_ranges: list[str] = []
    semantic_cell_refs: list[str] = []
    semantic_cell_values: list[str] = []
    sheet_names = list(sheets.keys())
    for sheet_name, frame in sheets.items():
        height, width = frame.shape
        used_ranges.append(used_range(sheet_name, width, height))
        columns = frame.get_columns()
        for col_index, column in enumerate(columns, start=1):
            col_name = column_name(col_index)
            for row_index, value in enumerate(column.to_list(), start=1):
                if value is None:
                    continue
                cell_count += 1
                ref = f"{sheet_name}!{col_name}{row_index}"
                semantic_cell_refs.append(ref)
                semantic_cell_values.append(f"{ref}\t{scalar_payload(value)}")
    return {
        "runnerVersion": pl.__version__,
        "sheetCount": len(sheet_names),
        "sheetNamesHash": hash_lines([f"{index}:{name}" for index, name in enumerate(sheet_names)]),
        "cellCount": cell_count,
        "physicalCellCount": cell_count,
        "formulaCount": 0,
        "usedRangeCount": len(used_ranges),
        "firstUsedRange": used_ranges[0] if used_ranges else None,
        "firstPhysicalUsedRange": used_ranges[0] if used_ranges else None,
        "usedRangesHash": hash_lines(used_ranges),
        "physicalUsedRangesHash": hash_lines(used_ranges),
        "semanticCellRefsHash": hash_lines(semantic_cell_refs),
        "semanticCellValuesHash": hash_lines(semantic_cell_values),
        "formulaTextHash": hash_lines([]),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--operation", choices=["read"], required=True)
    parser.add_argument("--file", required=True)
    parser.add_argument(
        "--engine", choices=["calamine", "xlsx2csv", "openpyxl"], default="calamine"
    )
    parser.add_argument("--repeat", type=int, default=1)
    parser.add_argument("--warmup", type=int, default=0)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    path = Path(args.file)
    for _ in range(max(0, args.warmup)):
        read_materialized(path, args.engine)
    samples: list[dict[str, float]] = []
    assertions: dict[str, str | int | bool | None] | None = None
    for _ in range(max(1, args.repeat)):
        before = memory_baseline()
        start = time.perf_counter()
        sheets = read_materialized(path, args.engine)
        duration_ms = (time.perf_counter() - start) * 1000
        assertions = read_assertions(sheets)
        assertions["runnerEngine"] = args.engine
        samples.append(sample_with_memory(duration_ms, before))
    payload: dict[str, Any] = {"assertions": assertions or {}, "samples": samples}
    if args.json:
        print(json.dumps(payload, separators=(",", ":")))
    else:
        print(json.dumps(payload, indent=2))


if __name__ == "__main__":
    main()
