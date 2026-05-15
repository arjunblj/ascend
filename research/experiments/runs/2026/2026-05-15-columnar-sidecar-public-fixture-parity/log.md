# Columnar Sidecar Public Fixture Parity

Date: 2026-05-15

## Question

Can the columnar sidecar proof run against a tracked public XLSX fixture and preserve checksum parity without claiming broad real-workbook performance?

## Hypothesis

Yes. A small public fixture can prove imported-workbook parity and report timing/payload evidence, while still rejecting broad performance claims until larger public workbook tables are covered.

## External sources checked

- DuckDB Excel import supports reading specific XLSX sheets and ranges, which is the external product contrast for query-shaped workbook ranges: https://duckdb.org/docs/lts/guides/file_formats/excel_import.html
- Apache Arrow's columnar format emphasizes adjacency for scans and implementation-owned mutation coordination, which matches Ascend's disposable sidecar boundary: https://arrow.apache.org/docs/format/Columnar.html
- Apache Arrow's overview describes the in-memory columnar format as a standardized representation for structured table-like datasets, but Ascend's current probe remains an internal numeric projection rather than Arrow ABI output: https://arrow.apache.org/overview/

## Why this matters to Ascend

The prior sidecar proof was synthetic. A public fixture run is stronger because it goes through real XLSX import and workbook model hydration before building the sidecar. It still must not become a product claim: the tested table is small, mixed values are only counted/skipped, and no production cache exists.

## Probe/implementation

Extended `fixtures/benchmarks/columnar-sidecar.ts`:

- added `runColumnarSidecarFixtureBenchmark` for public XLSX fixtures;
- added `--fixture`, `--sheet`, and `--range` CLI flags;
- reports source, fixture, sheet, open time, checksum parity, payload bytes, speedup, and generation invalidation;
- changes claim wording for fixture runs to "public-fixture evidence" and adds a do-not-promote boundary for broad performance claims.

Updated `fixtures/benchmarks/columnar-sidecar.test.ts` with a public fixture check against `fixtures/xlsx/poi/Tables.xlsx`, sheet `Exp1`.

## Results

Probe:

```bash
bun run fixtures/benchmarks/columnar-sidecar.ts --fixture fixtures/xlsx/poi/Tables.xlsx --sheet Exp1 --repeats 80 --claim-report --json
```

Observed:

- Source: fixture.
- Fixture: `fixtures/xlsx/poi/Tables.xlsx`.
- Sheet: `Exp1`.
- Range: `A1:J26`.
- Cells: 260.
- Populated count: 156.
- Numeric count: 119.
- Estimated sidecar payload bytes: 2,340.
- Matching generation valid: true.
- Next generation valid: false.
- End-to-end speedup in the probe run: 1.53x.
- Proof status: passed.

Validation:

```bash
bun test fixtures/benchmarks/columnar-sidecar.test.ts
bunx biome check fixtures/benchmarks/columnar-sidecar.ts fixtures/benchmarks/columnar-sidecar.test.ts
bunx tsc --build
```

## Confidence

Medium. Public fixture parity is now proven for one small imported workbook range. Confidence is still low for broad performance because the range is tiny and does not cover formulas, filters, hidden rows, merged cells, table totals, or query-backed tables.

## Fold-in decision

Fold into the performance proof harness. Keep public SDK/API/MCP surfaces and broad sidecar performance claims on the do-not-promote list.

## Next question

Can the performance loop find or create a larger public real-workbook table fixture that is safe to track, then prove parity, payload bytes, invalidation, and speedup without relying on private corpora?
