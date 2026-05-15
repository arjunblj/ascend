# Columnar Sidecar Claim Report

Date: 2026-05-15

## Question

Can the existing columnar sidecar benchmark emit a claim-safe proof report that includes memory/build cost and an honest no-promotion boundary?

## Hypothesis

Yes. Ascend can claim local evidence for disposable numeric sidecars on repeated dense scans, but only if the report names sidecar payload bytes, build cost, checksum parity, and the missing generation-invalidation proof.

## External sources checked

- Apache Arrow's columnar format lists data adjacency for sequential scans, constant-time random access, SIMD/vectorization friendliness, and notes that mutation coordination is left to implementations: https://arrow.apache.org/docs/format/Columnar.html
- Apache Arrow's overview frames the in-memory columnar format as a standardized representation for structured table-like datasets and highlights vectorized scanning: https://arrow.apache.org/overview/
- DuckDB Excel import supports selecting XLSX sheets and ranges with `read_xlsx`, making spreadsheet ranges query-shaped but not preservation-first workbook state: https://duckdb.org/docs/lts/guides/file_formats/excel_import.html
- DuckDB/Arrow integration describes reading Arrow objects through replacement scans, which supports the sidecar-as-derived-view mental model rather than replacing Ascend's workbook model: https://duckdb.org/2021/12/03/duck-arrow.html

## Why this matters to Ascend

Columnar sidecars are a plausible performance differentiator, but the claim is easy to overstate. Ascend should not imply it has an Arrow engine, DuckDB bridge, mixed-type table engine, or invalidation-safe production cache. The useful proof is narrower: a disposable, generation-stamped numeric projection can accelerate repeated scans while the workbook grid remains authoritative.

## Probe/implementation

Folded a claim-report mode into `fixtures/benchmarks/columnar-sidecar.ts`:

- records estimated sidecar payload bytes from `Float64Array` values plus validity bitmaps;
- emits structured and Markdown claim reports via `--claim-report` and `--claim-report --json`;
- includes allowed wording, proof status, boundary, kill criterion, and "do not promote yet" list;
- adds test coverage in `fixtures/benchmarks/columnar-sidecar.test.ts`.

## Results

Probe commands:

```bash
bun run fixtures/benchmarks/columnar-sidecar.ts --rows 20000 --cols 8 --repeats 20 --claim-report
bun run fixtures/benchmarks/columnar-sidecar.ts --rows 20000 --cols 8 --repeats 20 --claim-report --json
bun test fixtures/benchmarks/columnar-sidecar.test.ts
```

Observed from the Markdown run:

- Range: `A1:H20000`.
- Cells: 160,000.
- Repeats: 20.
- Estimated sidecar payload bytes: 1,440,000.
- Grid repeated scan: 25.414 ms.
- Sidecar build: 7.7 ms.
- Sidecar repeated scan: 3.907 ms.
- End-to-end speedup including build: 2.19x.
- Proof status: passed.

The JSON rerun also passed, with timing variation expected for local microbenchmarks.

## Confidence

Medium. The proof now includes build and memory cost and checksum parity, but it is still synthetic. Confidence should not rise until public real workbook tables and generation invalidation are tested.

## Fold-in decision

Fold into performance proof packaging only. Do not add a public sidecar surface, and do not promote product wording beyond "local evidence for disposable numeric sidecars on repeated dense scans."

## Next question

Can the sidecar benchmark run on public real workbook tables with checksum parity and report build/invalidation costs, or should the performance loop first define a generation invalidation API?
