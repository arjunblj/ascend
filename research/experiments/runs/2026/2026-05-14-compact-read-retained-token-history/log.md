# Compact Read Retained Token History

## Question

Can compact range reads patch from stale-but-retained `changeToken` snapshots instead of requiring the immediately previous token?

## Hypothesis

Yes. `SheetHandle.readWindowCompact()` already stores a per-window snapshot; extending that storage from one latest snapshot to a bounded ledger should move compact reads toward MVCC-style retained generations while preserving the existing fallback behavior.

## External sources checked

- PostgreSQL MVCC introduction: https://www.postgresql.org/docs/current/mvcc-intro.html
- RocksDB snapshots: https://github.com/facebook/rocksdb/wiki/Snapshot
- FoundationDB snapshot read discussion: https://forums.foundationdb.org/t/does-snapshot-provide-snapshot-level-isolation/301
- Automerge binary format heads: https://automerge.org/automerge-binary-format-spec/
- Automerge synchronization protocol: https://posit-dev.github.io/automerge-r/articles/sync-protocol.html

## Why this matters to Ascend

Agents and UI clients polling compact windows should not have to refresh just because one polling response was skipped. MVCC and LSM systems make older retained versions readable until compaction. A bounded retained-token ledger gives Ascend the same shape for compact reads: patch while retained, refresh when missing, invalid, stale outside the window, or expired.

## Probe/implementation

- Inspected `packages/sdk/src/sheet-handle.ts`: compact changedSince reads kept only the latest snapshot per window key.
- Inspected `packages/sdk/src/types.ts`: compact invalidation reasons did not include expiration.
- Inspected API/MCP forwarding and docs to confirm compact invalidation flows through public agent surfaces.
- Folded in a scoped SDK change:
  - added an internal retained snapshot ledger per compact window key;
  - retained the last 8 snapshots per key;
  - changed stale-but-retained tokens to diff against their retained base instead of invalidating;
  - added `base-token-expired` when the base token is older than the retained ledger;
  - preserved `base-snapshot-missing`, `base-token-stale`, and `base-token-invalid` behavior.
- Updated SDK docs/capability text for bounded token history.

## Results

Focused validation passed:

```bash
bun test packages/sdk/src/sdk.test.ts -t "compact changedSince"
bun test apps/api/src/server.test.ts -t "compact changedSince"
bun test apps/mcp/src/index.test.ts -t "compact changedSince"
bunx biome check packages/sdk/src/sheet-handle.ts packages/sdk/src/types.ts packages/sdk/src/sdk.test.ts packages/sdk/src/capabilities.ts docs/AGENT_API.md docs/AGENT_WORKFLOW.md research/experiments/index.md research/experiments/runs/2026/2026-05-14-compact-read-retained-token-history/log.md
bunx tsc --build
bun run test:changed
```

The updated test now verifies:

- immediate changedSince polling returns only changed cells;
- an older retained base token still patches without invalidation;
- a token older than the retained ledger returns a full window with `reason: "base-token-expired"`;
- invalid tokens still return `base-token-invalid`.

`bun run test:changed` expanded to the full suite and passed with 4973 tests, 1 skip, and 0 failures.

## Confidence

Medium-high. The change is internal to compact range reads and leaves API/MCP forwarding intact. It is still not the full shared viewport ledger for interactive sessions, but it is a real production step toward retained generation semantics.

## Fold-in decision

Folded into product/DX and performance loops for compact reads. Keep the larger `ViewportPatchLedger` as the next cross-surface design target.

## Next question

Can interactive session viewport patches share the compact retained-token ledger model without breaking existing session invalidation guarantees?
