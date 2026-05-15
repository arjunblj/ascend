# Columnar Sidecar Stress Fixture Scale

Date: 2026-05-15

## Question

Can the sidecar proof show scale on a tracked large public fixture while still refusing externally sourced real-workbook claims?

## Hypothesis

Yes. `fixtures/xlsx/stress/dense-100k.xlsx` can prove that the benchmark path handles a much larger imported workbook range with checksum parity, payload bytes, and generation invalidation. It should still be classified as stress-fixture evidence, not real-world table evidence.

## External sources checked

- DuckDB Excel import supports `read_xlsx` with sheet and range options, which reinforces why Ascend should compare against query-shaped range workloads without replacing workbook semantics: https://duckdb.org/docs/lts/guides/file_formats/excel_import.html
- Apache Arrow's columnar format documents sequential-scan-friendly layout and leaves mutation coordination to implementations, matching the sidecar proof boundary: https://arrow.apache.org/docs/format/Columnar.html
- Apache Arrow's overview frames columnar data as table-like in-memory analytics infrastructure; Ascend's current sidecar remains an internal numeric projection, not Arrow interchange output: https://arrow.apache.org/overview/

## Why this matters to Ascend

The previous public fixture proof used a tiny POI workbook. A 100k-row tracked stress workbook is stronger performance evidence and catches scale mistakes in sidecar payload accounting. It still does not prove real-world workbook behavior, so the claim must remain performance-loop evidence only.

## Probe/implementation

No new surface was added. The existing fixture-aware sidecar report was run against `fixtures/xlsx/stress/dense-100k.xlsx`, then report wording was tightened:

- public fixture evidence is no longer described as "small";
- the boundary now explicitly says this is not externally sourced real-world table proof;
- the do-not-promote list blocks broad real-workbook performance claims until externally sourced public tables are benchmarked.

## Results

Probe:

```bash
bun run fixtures/benchmarks/columnar-sidecar.ts --fixture fixtures/xlsx/stress/dense-100k.xlsx --repeats 20 --claim-report --json
```

Observed:

- Fixture: `fixtures/xlsx/stress/dense-100k.xlsx`.
- Sheet: `Data`.
- Range: `A1:E100000`.
- Cells: 500,000.
- Populated count: 500,000.
- Numeric count: 333,333.
- Estimated sidecar payload bytes: 4,500,000.
- Matching generation valid: true.
- Next generation valid: false.
- Grid repeated scan: 94.277 ms.
- Sidecar build: 11.475 ms.
- Sidecar repeated scan: 6.198 ms.
- End-to-end speedup including build: 5.33x.
- Proof status: passed.

Validation:

```bash
bun test fixtures/benchmarks/columnar-sidecar.test.ts
bunx biome check fixtures/benchmarks/columnar-sidecar.ts fixtures/benchmarks/columnar-sidecar.test.ts
bunx tsc --build
```

## Confidence

Medium-high for stress-fixture scale. Medium-low for real-workbook performance because the fixture is generated and does not cover workbook features such as filters, hidden rows, formulas, merged cells, table totals, or query-backed tables.

## Fold-in decision

Fold into the performance proof harness. Keep real-workbook sidecar product claims blocked until externally sourced public tables are benchmarked.

## Next question

Which externally sourced public workbook table can Ascend safely add or reference to prove sidecar parity and speedup without relying on private corpora?
