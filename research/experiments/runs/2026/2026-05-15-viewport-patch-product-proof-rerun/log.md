# Viewport Patch Product Proof Rerun

## Question

Can the retained viewport patch history claim be promoted to product-proof backed wording without adding CLI `changedSince` or implying CRDT/collaboration semantics?

## Hypothesis

Yes. The current evidence supports a guarded product claim: SDK interactive retained patches plus API/MCP compact read `changedSince` recovery. CLI should remain explicitly excluded.

## External sources checked

- PostgreSQL MVCC introduction: https://www.postgresql.org/docs/17/mvcc-intro.html
- SQLite isolation and WAL snapshot isolation: https://www.sqlite.org/isolation.html
- RocksDB snapshots: https://github.com/facebook/rocksdb/wiki/Snapshot
- RocksDB sequence number terminology: https://github.com/facebook/rocksdb/wiki/Terminology
- Automerge CRDT concepts: https://automerge.org/docs/reference/concepts/
- Yjs document updates: https://docs.yjs.dev/api/document-updates

## Why this matters to Ascend

Agents and UI clients need efficient refreshes after small workbook edits, but the claim must stay precise. Database snapshot systems justify bounded retained-token recovery; CRDT systems are a different category and should not be implied by Ascend's single-session viewport patch proof.

## Probe/implementation

- Ran `git status --short --branch`; no tracked production diff was present.
- Reran the tracked viewport proof harness:

```bash
bun run fixtures/benchmarks/viewport-patch-proof.ts
```

- Validated SDK/API/MCP surfaces:

```bash
bun test fixtures/benchmarks/viewport-patch-proof.test.ts
bun test packages/sdk/src/interactive-contract.test.ts -t "retained|viewport patch results expose invalidation|tokens from other sessions"
bun test apps/api/src/server.test.ts -t "compact changedSince"
bun test apps/mcp/src/index.test.ts -t "compact changedSince"
```

- Updated:
  - `research/experiments/syntheses/2026-05-viewport-patch-proof-report.md`
  - `research/experiments/syntheses/2026-05-release-claim-board.md`
  - `research/experiments/syntheses/2026-05-agent-context-contracts.md`

## Results

The proof rerun covered 7 SDK interactive cases:

- retained patch: `patch:A1`;
- skipped older retained token: `patch:A1=3`;
- invalid token: `base-token-invalid`;
- cross-session token: `base-snapshot-missing`;
- expired history: `base-token-expired`;
- projection change: `base-snapshot-missing`;
- metadata invalidation: `viewport-invalidated`.

All cases passed. Retained patch cases totaled 630 bytes.

Cross-surface validation passed:

- SDK interactive retained/invalidation tests: 4 tests.
- API compact `changedSince` tests: 2 tests.
- MCP compact `changedSince` tests: 3 tests.

## Confidence

High for guarded product-proof wording: SDK retained viewport patches plus API/MCP compact recovery. Low for cross-surface parity wording because CLI has no retained-patch or `changedSince` contract.

## Fold-in decision

Promote to product/performance proof packaging only with explicit surface wording: SDK retained patches; API/MCP recovery; CLI excluded. Do not add CLI `changedSince` in this loop.

## Next question

Can formula language-service primitives be made product-proof backed as a no-rename claim with cross-surface rejection evidence and latency checks, without adding edit-producing rename?
