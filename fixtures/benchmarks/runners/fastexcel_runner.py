#!/usr/bin/env python3
"""fastexcel runner for competitive real-workbook read benchmarks."""

from __future__ import annotations

import argparse
import hashlib
import json
import time
from pathlib import Path
from typing import Any

import fastexcel
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


def read_materialized(path: Path) -> tuple[Any, list[tuple[str, Any]]]:
    reader = fastexcel.read_excel(path)
    batches: list[tuple[str, Any]] = []
    for index, sheet_name in enumerate(reader.sheet_names):
        sheet = reader.load_sheet(index, header_row=None)
        batches.append((sheet_name, sheet.to_arrow()))
    return reader, batches


def read_assertions(reader: Any, batches: list[tuple[str, Any]]) -> dict[str, str | int | bool | None]:
    cell_count = 0
    used_ranges: list[str] = []
    semantic_cell_refs: list[str] = []
    semantic_cell_values: list[str] = []
    for sheet_name, batch in batches:
        width = int(batch.num_columns or 0)
        height = int(batch.num_rows or 0)
        cell_count += sum(len(column) - column.null_count for column in batch.columns)
        for col_index, column in enumerate(batch.columns, start=1):
            col_name = column_name(col_index)
            for row_index in range(len(column)):
                value = column[row_index].as_py()
                if value is not None:
                    ref = f"{sheet_name}!{col_name}{row_index + 1}"
                    semantic_cell_refs.append(ref)
                    semantic_cell_values.append(f"{ref}\t{scalar_payload(value)}")
        used_ranges.append(used_range(sheet_name, width, height))
    return {
        "runnerVersion": getattr(fastexcel, "__version__", "unknown"),
        "sheetCount": len(reader.sheet_names),
        "sheetNamesHash": hash_lines([f"{index}:{name}" for index, name in enumerate(reader.sheet_names)]),
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
        "readCommentCount": 0,
        "readHyperlinkCount": 0,
        "readDataValidationCount": 0,
        "readConditionalFormatCount": 0,
        "readDefinedNameCount": 0,
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
        before = memory_baseline()
        start = time.perf_counter()
        reader, batches = read_materialized(path)
        duration_ms = (time.perf_counter() - start) * 1000
        assertions = read_assertions(reader, batches)
        samples.append(sample_with_memory(duration_ms, before))
    payload: dict[str, Any] = {"assertions": assertions or {}, "samples": samples}
    if args.json:
        print(json.dumps(payload, separators=(",", ":")))
    else:
        print(json.dumps(payload, indent=2))


if __name__ == "__main__":
    main()
