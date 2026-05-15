# Columnar Sidecar Generation Invalidation

Date: 2026-05-15

## Question

Can the columnar sidecar proof make invalidation explicit with generation matching before any production cache is promoted?

## Hypothesis

Yes. A sidecar stamped with workbook generation can be treated as usable only when the current generation matches. That is not a cache manager, but it proves the minimum invalidation predicate the performance loop must preserve.

## External sources checked

- PostgreSQL MVCC documentation frames reads as seeing a snapshot of data as of a point in time, independent of later writes: https://www.postgresql.org/docs/18/mvcc-intro.html
- RocksDB snapshots are point-in-time read-only views, which is the right mental model for derived sidecars that should not silently survive writes: https://github.com/facebook/rocksdb/wiki/Snapshot
- Apache Arrow's columnar format documentation states that mutation coordination is left to implementations, supporting Ascend's decision to keep generation invalidation outside the columnar layout itself: https://arrow.apache.org/docs/format/Columnar.html

## Why this matters to Ascend

The columnar sidecar claim should not become "fast but stale." A generation predicate is the smallest proof that sidecars are derived from a specific workbook state and must be rejected after mutation. This keeps the workbook grid authoritative.

## Probe/implementation

Folded generation validation into `fixtures/benchmarks/columnar-sidecar.ts`:

- added `isColumnarSidecarCurrent(sidecar, currentGeneration)`;
- added generation invalidation fields to the claim report;
- rendered matching and next-generation validity in Markdown;
- added test assertions that a sidecar is valid for its own generation and invalid for the next generation.

## Results

Validation:

```bash
bun run fixtures/benchmarks/columnar-sidecar.ts --rows 20000 --cols 8 --repeats 20 --claim-report --json
bun test fixtures/benchmarks/columnar-sidecar.test.ts
bunx biome check fixtures/benchmarks/columnar-sidecar.ts fixtures/benchmarks/columnar-sidecar.test.ts
bunx tsc --build
```

Observed:

- `generationInvalidation.sidecarGeneration=1`.
- `matchingGenerationValid=true`.
- `nextGenerationValid=false`.
- End-to-end sidecar speedup in the validation run: 2.53x.
- Estimated sidecar payload bytes: 1,440,000.

## Confidence

Medium. The predicate is simple and tested, but it is not yet wired into workbook mutation lifecycle or real workbook workloads.

## Fold-in decision

Fold into the performance proof harness. Do not promote a production sidecar cache until this predicate is owned by a generation-aware cache/invalidation API.

## Next question

Can the performance loop run sidecar scans over public real workbook tables and prove checksum parity, speedup, payload bytes, and generation invalidation together?
