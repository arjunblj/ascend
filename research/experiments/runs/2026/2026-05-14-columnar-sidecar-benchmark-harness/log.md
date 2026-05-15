# Columnar Sidecar Benchmark Harness

## Question

Can the columnar sidecar range-scan probe be folded into a reusable benchmark harness that measures range/table scan acceleration without changing workbook semantics?

## Hypothesis

Yes. The original research probe showed a large synthetic speedup, but it was not reusable by benchmark loops. A small benchmark module can keep the sidecar experimental while making repeated measurements and correctness checks easy to run.

## External sources checked

- DuckDB Excel extension: https://duckdb.org/docs/lts/core_extensions/excel.html
- DuckDB SQL on Apache Arrow: https://duckdb.org/docs/current/guides/python/sql_on_arrow.html
- DuckDB Arrow integration post: https://duckdb.org/2021/12/03/duck-arrow.html
- DuckDB replacement scans: https://duckdb.org/docs/stable/clients/c/replacement_scans.html
- Apache Arrow JS docs: https://arrow.apache.org/docs/4.0/js/index.html
- Apache Arrow columnar format: https://arrow.apache.org/docs/format/Columnar.html

## Why this matters to Ascend

Ascend should preserve spreadsheet semantics while becoming fast at database-shaped questions over ranges and tables. DuckDB and Arrow show that derived columnar views are a practical analytics boundary. The important Ascend-specific constraint is that the sidecar must remain disposable evidence over workbook state, not a replacement model.

## Probe/implementation

- Inspected the original ignored probe at `research/experiments/runs/2026/2026-05-14-columnar-sidecar-range-scans/probes/columnar-sidecar-scan.ts`.
- Inspected existing benchmark conventions under `fixtures/benchmarks`.
- Added `fixtures/benchmarks/columnar-sidecar.ts`:
  - builds deterministic dense numeric sheets;
  - builds numeric column sidecars with `Float64Array` values and `Uint8Array` validity bitmaps;
  - records range, generation, populated count, numeric count, checksum, build time, scan time, repeated-scan speedup, and end-to-end speedup;
  - exposes `runColumnarSidecarBenchmark()`, `buildNumericColumnSidecar()`, and `sumSidecarColumn()`;
  - can run directly with `--rows`, `--cols`, `--repeats`, and `--generation`.
- Added `fixtures/benchmarks/columnar-sidecar.test.ts` to validate checksum parity and mixed-value validity metadata without adding a timing threshold.

## Results

Focused validation passed:

```bash
bun test fixtures/benchmarks/columnar-sidecar.test.ts
bun run fixtures/benchmarks/columnar-sidecar.ts --rows 50000 --cols 8 --repeats 40
bunx biome check fixtures/benchmarks/columnar-sidecar.ts fixtures/benchmarks/columnar-sidecar.test.ts research/experiments/index.md research/experiments/runs/2026/2026-05-14-columnar-sidecar-benchmark-harness/log.md
bunx tsc --build
bun run test:changed
```

Harness output for 50,000 x 8 x 40:

| Metric | Value |
| --- | ---: |
| Cells | 400,000 |
| Workbook build | 160.028 ms |
| Sidecar build | 21.784 ms |
| Grid repeated scan | 978.997 ms |
| Sidecar repeated scan | 29.108 ms |
| Sidecar end-to-end | 50.892 ms |
| Repeated scan speedup | 33.63x |
| End-to-end speedup | 19.24x |

`bun run test:changed` expanded to the full suite and passed with 4971 tests, 1 skip, and 0 failures.

## Confidence

Medium-high for the benchmark harness and numeric sidecar correctness. Medium for product semantics: hidden rows, filtered tables, merged cells, rich text, formulas, and mutation invalidation remain out of scope.

## Fold-in decision

Folded into the performance loop as a reusable benchmark harness. No workbook runtime behavior or public SDK API changed.

## Next question

Can UI patch streams use MVCC-style generation snapshots to make compact reads and TUI viewport updates safer under concurrent agent edits?
