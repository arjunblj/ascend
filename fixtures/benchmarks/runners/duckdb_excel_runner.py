#!/usr/bin/env python3
"""DuckDB excel-extension runner for generated XLSX benchmark profiles."""

from __future__ import annotations

import argparse
import hashlib
import json
import tempfile
import time
import zipfile
from decimal import Decimal
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

import duckdb
from generated_xlsx_validation import expected_ordered_values_hash, generated_write_assertions
from memory_metrics import memory_baseline, sample_with_memory


SUPPORTED_WRITE_WORKLOADS = {
    "dense-values",
    "mixed-closedxml-10text-5number",
    "sparse-wide",
    "styles-heavy",
}


def strip_namespace(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def column_name(index: int) -> str:
    name = ""
    current = index
    while current > 0:
        current, rem = divmod(current - 1, 26)
        name = chr(65 + rem) + name
    return name or "A"


def column_index(name: str) -> int:
    value = 0
    for char in name:
        if char < "A" or char > "Z":
            break
        value = value * 26 + (ord(char) - 64)
    return value


def hash_lines(lines: list[str]) -> str:
    digest = hashlib.sha256()
    for line in sorted(lines):
        digest.update(f"{len(line)}:".encode("utf-8"))
        digest.update(line.encode("utf-8"))
        digest.update(b"\n")
    return digest.hexdigest()


def sql_string(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def workbook_first_sheet(path: Path) -> tuple[str, str]:
    with zipfile.ZipFile(path) as archive:
        workbook_xml = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        first_sheet = next(
            element for element in workbook_xml.iter() if strip_namespace(element.tag) == "sheet"
        )
        sheet_name = first_sheet.attrib.get("name", "Sheet1")
        rel_id = first_sheet.attrib.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        if rel_id:
            rels_xml = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            for relationship in rels_xml.iter():
                if strip_namespace(relationship.tag) != "Relationship":
                    continue
                if relationship.attrib.get("Id") != rel_id:
                    continue
                target = relationship.attrib.get("Target", "worksheets/sheet1.xml")
                if target.startswith("/"):
                    return sheet_name, target.lstrip("/")
                return sheet_name, str(Path("xl") / target)
    return sheet_name, "xl/worksheets/sheet1.xml"


def worksheet_shape(path: Path, worksheet_path: str) -> tuple[int, int]:
    with zipfile.ZipFile(path) as archive:
        with archive.open(worksheet_path) as handle:
            max_row = 0
            max_col = 0
            for event, element in ElementTree.iterparse(handle, events=("start",)):
                tag = strip_namespace(element.tag)
                if tag == "dimension":
                    ref = element.attrib.get("ref", "")
                    if ":" in ref:
                        ref = ref.rsplit(":", 1)[-1]
                    col = "".join(char for char in ref if char.isalpha())
                    row = "".join(char for char in ref if char.isdigit())
                    if col and row:
                        return int(row), column_index(col)
                elif tag == "c":
                    ref = element.attrib.get("r", "")
                    col = "".join(char for char in ref if char.isalpha())
                    row = "".join(char for char in ref if char.isdigit())
                    if col and row:
                        max_row = max(max_row, int(row))
                        max_col = max(max_col, column_index(col))
    return max_row, max_col


def prepare_connection() -> duckdb.DuckDBPyConnection:
    connection = duckdb.connect(database=":memory:")
    try:
        connection.execute("LOAD excel")
    except Exception:
        connection.execute("INSTALL excel")
        connection.execute("LOAD excel")
    return connection


def canonical_number(value: int | float | Decimal) -> str:
    number = float(value)
    if number == 0:
        return "0"
    if number.is_integer():
        return str(int(number))
    return format(number, ".15g")


def scalar_payload(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return f"b:{str(value).lower()}"
    if isinstance(value, (int, float, Decimal)):
        return f"n:{canonical_number(value)}"
    return f"s:{value}"


def used_range(sheet_name: str, rows: int, cols: int) -> str:
    if rows <= 0 or cols <= 0:
        return f"{sheet_name}!empty"
    return f"{sheet_name}!A1:{column_name(cols)}{rows}"


def read_materialized(
    connection: duckdb.DuckDBPyConnection,
    path: Path,
) -> tuple[str, int, int, list[tuple[Any, ...]]]:
    sheet_name, worksheet_path = workbook_first_sheet(path)
    rows, cols = worksheet_shape(path, worksheet_path)
    range_ref = f"A1:{column_name(cols)}{rows}" if rows > 0 and cols > 0 else "A1:A1"
    query = """
        SELECT *
        FROM read_xlsx(
            ?,
            header = false,
            sheet = ?,
            range = ?,
            stop_at_empty = false,
            ignore_errors = false
        )
    """
    result = connection.execute(query, [str(path), sheet_name, range_ref]).fetchall()
    return sheet_name, rows, cols, result


def read_assertions(
    sheet_name: str,
    rows: int,
    cols: int,
    values: list[tuple[Any, ...]],
) -> dict[str, str | int | bool | None]:
    semantic_cell_refs: list[str] = []
    semantic_cell_values: list[str] = []
    column_names = [column_name(col + 1) for col in range(cols)]
    for row_index, row in enumerate(values[:rows], start=1):
        for col_index, value in enumerate(row[:cols], start=1):
            payload = scalar_payload(value)
            if payload is None:
                continue
            ref = f"{sheet_name}!{column_names[col_index - 1]}{row_index}"
            semantic_cell_refs.append(ref)
            semantic_cell_values.append(f"{ref}\t{payload}")
    range_label = used_range(sheet_name, rows, cols)
    return {
        "runnerVersion": duckdb.__version__,
        "sheetCount": 1,
        "sheetNamesHash": hash_lines([f"0:{sheet_name}"]),
        "cellCount": len(semantic_cell_values),
        "physicalCellCount": len(semantic_cell_values),
        "formulaCount": 0,
        "usedRangeCount": 1 if rows > 0 and cols > 0 else 0,
        "firstUsedRange": range_label,
        "firstPhysicalUsedRange": range_label,
        "usedRangesHash": hash_lines([range_label] if rows > 0 and cols > 0 else []),
        "physicalUsedRangesHash": hash_lines([range_label] if rows > 0 and cols > 0 else []),
        "semanticCellRefsHash": hash_lines(semantic_cell_refs),
        "semanticCellValuesHash": hash_lines(semantic_cell_values),
        "formulaTextHash": hash_lines([]),
    }


def workload_expression(workload: str, col: int, cols: int) -> str:
    if workload == "dense-values":
        return f"row_id * {cols} + {col}"
    if workload == "styles-heavy":
        return f"(row_id + 1) * {col + 1}"
    if workload == "mixed-closedxml-10text-5number":
        return "'Hello world'" if col < 10 else str(col - 10)
    if workload == "sparse-wide":
        if col == 0:
            return "row_id"
        if col == cols - 1:
            return f"'edge-' || CAST(row_id AS VARCHAR) || '-{cols}'"
        return (
            f"CASE WHEN ((row_id * 31 + {col} * 17) % 97) = 0 "
            f"THEN row_id * {cols} + {col} ELSE NULL END"
        )
    supported = ", ".join(sorted(SUPPORTED_WRITE_WORKLOADS))
    raise ValueError(f'DuckDB Excel writer supports only {supported}; got "{workload}"')


def write_query(workload: str, rows: int, cols: int) -> str:
    expressions = [
        f"{workload_expression(workload, col, cols)} AS c{col + 1}" for col in range(cols)
    ]
    return f"SELECT {', '.join(expressions)} FROM range(0, {rows}) AS source(row_id)"


def write_workbook(
    connection: duckdb.DuckDBPyConnection,
    workload: str,
    rows: int,
    cols: int,
) -> bytes:
    if rows < 0 or cols < 0:
        raise ValueError("--rows and --cols must be non-negative")
    with tempfile.TemporaryDirectory() as directory:
        output_path = Path(directory) / "duckdb-output.xlsx"
        query = write_query(workload, rows, cols)
        connection.execute(
            f"COPY ({query}) TO {sql_string(str(output_path))} "
            "WITH (FORMAT xlsx, SHEET 'Data', HEADER false)"
        )
        return output_path.read_bytes()


def write_assertions(
    data: bytes,
    workload: str,
    rows: int,
    cols: int,
    expected_ordered_hash: str | None,
) -> dict[str, str | int | bool | None]:
    assertions = generated_write_assertions(
        data,
        workload=workload,
        rows=rows,
        cols=cols,
        runner_version=duckdb.__version__,
        expected_ordered_hash=expected_ordered_hash,
    )
    assertions["duckdbExcelExtension"] = True
    return assertions


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--operation", choices=["read", "write"], required=True)
    parser.add_argument("--file")
    parser.add_argument("--rows", type=int)
    parser.add_argument("--cols", type=int)
    parser.add_argument("--workload", default="dense-values")
    parser.add_argument("--repeat", type=int, default=1)
    parser.add_argument("--warmup", type=int, default=0)
    parser.add_argument("--validation-mode", choices=["each", "final"], default="each")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    connection = prepare_connection()
    try:
        samples: list[dict[str, float]] = []
        assertions: dict[str, str | int | bool | None] | None = None
        if args.operation == "read":
            if not args.file:
                raise ValueError("--file is required for read")
            path = Path(args.file)
            for _ in range(max(0, args.warmup)):
                read_materialized(connection, path)
            for _ in range(max(1, args.repeat)):
                before = memory_baseline()
                start = time.perf_counter()
                sheet_name, rows, cols, values = read_materialized(connection, path)
                duration_ms = (time.perf_counter() - start) * 1000
                assertions = read_assertions(sheet_name, rows, cols, values)
                samples.append(sample_with_memory(duration_ms, before))
        else:
            if args.rows is None or args.cols is None:
                raise ValueError("--rows and --cols are required for write")
            for _ in range(max(0, args.warmup)):
                write_workbook(connection, args.workload, args.rows, args.cols)
            expected_ordered_hash = expected_ordered_values_hash(
                args.workload, args.rows, args.cols
            )
            data: bytes | None = None
            for _ in range(max(1, args.repeat)):
                before = memory_baseline()
                start = time.perf_counter()
                data = write_workbook(connection, args.workload, args.rows, args.cols)
                duration_ms = (time.perf_counter() - start) * 1000
                if args.validation_mode == "each":
                    assertions = write_assertions(
                        data, args.workload, args.rows, args.cols, expected_ordered_hash
                    )
                samples.append(sample_with_memory(duration_ms, before))
            if args.validation_mode == "final":
                if data is None:
                    raise RuntimeError("no workbook bytes were produced")
                assertions = write_assertions(
                    data, args.workload, args.rows, args.cols, expected_ordered_hash
                )
            if assertions is not None:
                assertions["validationMode"] = args.validation_mode
                assertions["validationSamples"] = args.repeat if args.validation_mode == "each" else 1

        payload: dict[str, Any] = {"assertions": assertions or {}, "samples": samples}
        if args.json:
            print(json.dumps(payload, separators=(",", ":")))
        else:
            print(json.dumps(payload, indent=2))
    finally:
        connection.close()


if __name__ == "__main__":
    main()
