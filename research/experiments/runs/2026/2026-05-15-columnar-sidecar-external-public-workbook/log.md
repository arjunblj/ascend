# Columnar Sidecar External Public Workbook

Date: 2026-05-15

## Question

Can the sidecar proof run against an externally sourced public workbook already tracked in `fixtures/xlsx/external` and preserve checksum parity?

## Hypothesis

Yes. The SEC money market fund statistics workbook has real public reporting structure and numeric sheets. A sidecar run over one dense imported sheet should prove parity and report performance while still blocking broad real-world claims.

## External sources checked

- `fixtures/xlsx/external/download.sh` records the SEC source URL for `mmf-statistics-supporting-data-2022-02xlsx`: https://www.sec.gov/file/mmf-statistics-supporting-data-2022-02xlsx
- U.S. Census Construction Spending historical data publishes XLSX files, matching the repository's external Census fixture source family: https://www.census.gov/construction/c30/historical_data.html
- NICE transparency of spend publishes monthly Excel files, matching the repository's external UK spend fixture source family: https://www.nice.org.uk/about-us/policies-procedures-and-reports/transparency-of-spend
- DuckDB Excel import supports sheet and range selection for XLSX data, useful competitor contrast for query-shaped workbook ranges: https://duckdb.org/docs/lts/guides/file_formats/excel_import.html
- Apache Arrow's columnar format documents scan-friendly column layout while leaving mutation coordination to implementations: https://arrow.apache.org/docs/format/Columnar.html

## Why this matters to Ascend

Columnar sidecars only become product-relevant if they work on workbook data that came from the outside world, not only generated stress files. The SEC workbook is externally sourced, public, and multi-sheet. This gives stronger evidence while keeping the workbook grid authoritative.

## Probe/implementation

Extended `fixtures/benchmarks/columnar-sidecar.ts` so fixture runs under `fixtures/xlsx/external/` are classified as `external-fixture` and receive stricter claim wording:

- allowed claim says "externally sourced public-workbook evidence";
- boundary says "one numeric/date-like imported range";
- do-not-promote list blocks broad real-workbook claims until multiple larger and structurally diverse external public workbooks are benchmarked.

Added `fixtures/benchmarks/columnar-sidecar.test.ts` coverage for `fixtures/xlsx/external/sec-mmf-statistics-2022-02.xlsx`, sheet `Table 9`.

## Results

Probe:

```bash
bun run fixtures/benchmarks/columnar-sidecar.ts --fixture fixtures/xlsx/external/sec-mmf-statistics-2022-02.xlsx --sheet 'Table 9' --repeats 40 --claim-report --json
```

Observed:

- Source: external fixture.
- Fixture: `fixtures/xlsx/external/sec-mmf-statistics-2022-02.xlsx`.
- Sheet: `Table 9`.
- Range: `A1:J115`.
- Cells: 1,150.
- Populated count: 1,141.
- Numeric count: 1,017.
- Estimated sidecar payload bytes: 10,350.
- Matching generation valid: true.
- Next generation valid: false.
- Grid repeated scan: 1.215 ms.
- Sidecar build: 0.517 ms.
- Sidecar repeated scan: 0.324 ms.
- End-to-end speedup including build: 1.45x.
- Proof status: passed.

## Confidence

Medium. This is the first external public workbook proof and it covers import parity, payload bytes, and generation invalidation for one numeric/date-like range. It is not enough for broad real-world performance wording.

## Fold-in decision

Fold into performance proof packaging only. Continue blocking public SDK/API/MCP sidecar surfaces and broad real-world performance claims.

## Next question

Can the proof cover multiple external public workbook sources in one report, including SEC, Census, and UK spend data, while preserving checksum parity and honest per-fixture boundaries?
