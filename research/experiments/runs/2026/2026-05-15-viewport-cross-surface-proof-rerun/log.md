# Viewport Cross-Surface Proof Rerun

## Question

Can retained viewport patch history be rerun and tied to current API/MCP compact read evidence without adding a new public surface?

## Hypothesis

Yes, with guarded wording. SDK interactive sessions prove retained patches; API and MCP compact reads prove `changedSince` invalidation recovery. CLI parity remains missing, so the release claim should not overstate cross-surface coverage.

## External sources checked

- PostgreSQL MVCC introduction: https://www.postgresql.org/docs/17/mvcc-intro.html
- SQLite isolation and WAL snapshot isolation: https://www.sqlite.org/isolation.html
- RocksDB snapshots: https://github.com/facebook/rocksdb/wiki/Snapshot
- Automerge CRDT concepts: https://automerge.org/docs/reference/concepts/
- Yjs document updates/state-vector model: https://docs.yjs.dev/api/document-updates

## Why this matters to Ascend

The release claim board lists retained viewport patch history as "needs one more fold-in." The useful product claim is narrow: agents can use bounded tokens when available and receive explicit refresh instructions when not. That is valuable for polling and retry loops, but it should not imply collaboration or unlimited history.

## Probe/implementation

- Reran the tracked SDK viewport proof harness.
- Reran targeted API and MCP compact `changedSince` tests.
- Created `research/experiments/syntheses/2026-05-viewport-patch-proof-report.md`.
- Updated the release claim board and experiment index.
- Did not change production code or add CLI/API/MCP surfaces.

## Results

SDK proof:

- `retained-patch`: passed, 315 patch bytes.
- `skipped-token-retained`: passed, 315 patch bytes.
- `invalid-token`: `base-token-invalid`.
- `cross-session-token`: `base-snapshot-missing`.
- `expired-history`: `base-token-expired`.
- `projection-change`: `base-snapshot-missing`.
- `metadata-invalidation`: `viewport-invalidated`.

Validation:

```bash
bun run fixtures/benchmarks/viewport-patch-proof.ts
bun test fixtures/benchmarks/viewport-patch-proof.test.ts
bun test apps/api/src/server.test.ts -t "compact changedSince"
bun test apps/mcp/src/index.test.ts -t "compact changedSince"
```

Observed result: SDK proof passed all seven cases; API passed two targeted tests; MCP passed three targeted tests.

## Confidence

High for SDK retained patch and API/MCP invalidation recovery. Medium for product wording because CLI coverage is absent and API/MCP compact reads are not identical to SDK interactive retained patches.

## Fold-in decision

Promote to topic synthesis and guarded product/performance handoff. Do not add new surfaces in this block.

## Next question

Should the product loop add a CLI `changedSince` compact read contract, or should the release claim explicitly exclude CLI and focus on SDK/API/MCP?
