#!/usr/bin/env python3
"""python-calamine runner for competitive real-workbook read benchmarks."""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import time
from datetime import date, datetime, time as datetime_time
from pathlib import Path
from typing import Any

import python_calamine
from memory_metrics import sample_with_memory

RUNNER_VERSION = importlib.metadata.version("python-calamine")


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


def canonical_number(value: Any) -> str:
    number = float(value)
    if number == 0:
        return "0"
    if number.is_integer():
        return str(int(number))
    return format(number, ".15g")


def scalar_payload(value: Any) -> str:
    if value is None or value == "":
        return "empty"
    if isinstance(value, bool):
        return f"b:{'true' if value else 'false'}"
    if isinstance(value, (int, float)):
        return f"n:{canonical_number(value)}"
    if isinstance(value, (date, datetime, datetime_time)):
        return f"s:{value.isoformat()}"
    return f"s:{value}"


def used_range(sheet_name: str, coords: list[tuple[int, int]]) -> str:
    if not coords:
        return f"{sheet_name}!empty"
    min_row = min(row for row, _col in coords)
    min_col = min(col for _row, col in coords)
    max_row = max(row for row, _col in coords)
    max_col = max(col for _row, col in coords)
    return f"{sheet_name}!{column_name(min_col)}{min_row}:{column_name(max_col)}{max_row}"


def read_materialized(path: Path) -> Any:
    workbook = python_calamine.load_workbook(path)
    try:
        sheets: list[tuple[str, list[list[Any]]]] = []
        for sheet_name in workbook.sheet_names:
            sheet = workbook.get_sheet_by_name(sheet_name)
            sheets.append((sheet_name, sheet.to_python()))
        return sheets
    finally:
        workbook.close()


def read_assertions(sheets: list[tuple[str, list[list[Any]]]]) -> dict[str, str | int | bool | None]:
    cell_count = 0
    used_ranges: list[str] = []
    semantic_cell_refs: list[str] = []
    semantic_cell_values: list[str] = []
    sheet_names = [sheet_name for sheet_name, _rows in sheets]
    for sheet_name, rows in sheets:
        semantic_coords: list[tuple[int, int]] = []
        for row_index, row in enumerate(rows, start=1):
            for col_index, value in enumerate(row, start=1):
                payload = scalar_payload(value)
                if payload == "empty":
                    continue
                cell_count += 1
                semantic_coords.append((row_index, col_index))
                ref = f"{sheet_name}!{column_name(col_index)}{row_index}"
                semantic_cell_refs.append(ref)
                semantic_cell_values.append(f"{ref}\t{payload}")
        used_ranges.append(used_range(sheet_name, semantic_coords))
    return {
        "runnerVersion": RUNNER_VERSION,
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
    parser.add_argument("--repeat", type=int, default=1)
    parser.add_argument("--warmup", type=int, default=0)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    path = Path(args.file)
    for _ in range(max(0, args.warmup)):
        read_materialized(path)
    samples: list[dict[str, float]] = []
    assertions: dict[str, str | int | bool | None] | None = None
    for _ in range(max(1, args.repeat)):
        start = time.perf_counter()
        sheets = read_materialized(path)
        duration_ms = (time.perf_counter() - start) * 1000
        assertions = read_assertions(sheets)
        samples.append(sample_with_memory(duration_ms))
    payload: dict[str, Any] = {"assertions": assertions or {}, "samples": samples}
    if args.json:
        print(json.dumps(payload, separators=(",", ":")))
    else:
        print(json.dumps(payload, indent=2))


if __name__ == "__main__":
    main()
