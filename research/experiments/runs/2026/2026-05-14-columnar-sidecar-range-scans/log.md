# Columnar Sidecar Range Scans

## Question

Can Arrow/DuckDB-style columnar sidecars accelerate table and range scans without replacing workbook semantics?

## Hypothesis

Yes. Ascend's canonical workbook should remain preservation-first and cell-addressable, but repeated analytical scans over a stable rectangular range can use a derived, generation-stamped, columnar sidecar. The sidecar should be disposable and invalidated by mutations, not a replacement storage engine.

## External sources checked

- [DuckDB Excel extension](https://duckdb.org/docs/lts/core_extensions/excel): DuckDB can read XLSX ranges as tables, with explicit `range`, `header`, and type-inference controls.
- [DuckDB export to Apache Arrow](https://duckdb.org/docs/current/guides/python/export_arrow): DuckDB can query Arrow tables and export query results as Arrow tables or record batch readers.
- [DuckDB relational API](https://duckdb.org/docs/stable/clients/python/relational_api): DuckDB relations are lazy and can be created from Arrow objects, which supports a derived-view mental model.
- [Apache Arrow documentation](https://arrow.apache.org/docs/): Arrow is a universal columnar format for fast data interchange and in-memory analytics.
- [Arrow columnar format](https://arrow.apache.org/docs/format/Columnar.html): Arrow's layout favors sequential scans, constant-time random access, SIMD/vectorization, and zero-copy-friendly memory.

## Why this matters to Ascend

Spreadsheet users and agents often ask database-shaped questions of spreadsheet-shaped data: summarize a table, preview a filtered column, compute range statistics, inspect missing values, or answer a natural-language question about a rectangular region. Today Ascend can preserve workbook semantics and stream/aggregate ranges, but analytic workloads pay row/cell traversal costs repeatedly.

A columnar sidecar could make Ascend feel database-native for scans while keeping XLSX semantics authoritative. The key is treating the sidecar as an indexed view over workbook state, with provenance and invalidation.

## Probe/implementation

Inspected local implementation:

- `packages/core/src/sparse-grid.ts` stores cells in chunked sparse/dense grid chunks and already has optimized methods such as `forEachValueInRange`, `iterateRowsInRange`, and `aggregateNumericInRange`.
- `packages/sdk/src/sheet-handle.ts` implements `aggregateRange` by visiting materialized cells in a parsed range and producing status-bar-style counts and numeric stats.
- `packages/sdk/src/range-aggregates.test.ts` confirms sparse selections are summarized without materializing empty cells.
- `packages/engine/src/calc.ts` uses range scan helpers for aggregate formula evaluation and keeps aggregate caches for calculation paths.

Added ignored probe `research/experiments/runs/2026/2026-05-14-columnar-sidecar-range-scans/probes/columnar-sidecar-scan.ts`. It:

1. Builds a synthetic 50,000 row x 8 column numeric workbook using `SparseGrid.setPlainNumberSpan`.
2. Builds a derived numeric sidecar with one `Float64Array` and one validity bitmap per column.
3. Compares repeated per-column sums through `SparseGrid.aggregateNumericInRange` against repeated sidecar scans.
4. Verifies both paths produce the same checksum.

Validation command:

```bash
bun run research/experiments/runs/2026/2026-05-14-columnar-sidecar-range-scans/probes/columnar-sidecar-scan.ts
```

## Results

Probe output:

| Metric | Value |
| --- | ---: |
| Rows | 50,000 |
| Columns | 8 |
| Cells | 400,000 |
| Repeated scan count | 40 |
| Workbook build | 115.321 ms |
| Sidecar build | 25.135 ms |
| Grid repeated scans | 288.951 ms |
| Sidecar repeated scans | 26.587 ms |
| Sidecar build + scans | 51.722 ms |
| Repeated scan speedup | 10.87x |
| End-to-end speedup including sidecar build | 5.59x |

The existing grid is already strong for sparse workbook semantics, but repeated analytical scans over dense rectangular numeric ranges benefit substantially from columnar layout. The observed sidecar build cost is low enough that it pays back quickly when agents ask multiple questions about the same table or viewport.

Important boundary: the probe only tested numeric dense ranges. Mixed values, formulas, rich strings, hidden rows, filters, table totals, merged cells, and mutation invalidation need more design before production.

## Confidence

Medium. The speedup is concrete and the external model fits Ascend's workload, but this is a synthetic probe. Confidence should rise after testing real workbook tables and after adding generation-aware invalidation semantics.

## Fold-in decision

Promote to performance loop.

Recommended fold-in:

1. Add a private experimental `RangeScanSidecar` in the performance loop, not public SDK API.
2. Key each sidecar by workbook id, sheet id, range, value projection, and workbook generation.
3. Support numeric/date columns first, with validity bitmaps and a fallback to grid scans for mixed columns.
4. Emit proof metadata: source range, generation, populated count, numeric count, checksum, build time, and invalidation reason.
5. Benchmark against real tables and agent-view workloads before exposing anything to product/DX loops.

Do not fold into production in this research loop. The next performance loop should own invalidation and real-workbook benchmarks.

## Next question

Can UI patch streams use database MVCC/LSM ideas for generation-aware snapshots?
