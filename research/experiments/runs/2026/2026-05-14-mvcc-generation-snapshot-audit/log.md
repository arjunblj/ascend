# MVCC Generation Snapshot Audit

## Question

Can MVCC-style generation snapshot metadata be folded into compact reads and TUI viewport updates now?

## Hypothesis

Only as a cross-surface protocol change. Compact reads already expose `changeToken` and invalidations, and interactive sessions have a richer viewport patch protocol. Adding generation metadata to one surface would be easy, but doing it without aligning SDK, API, MCP, and TUI/session semantics would create two incompatible patch models.

## External sources checked

- PostgreSQL MVCC introduction: https://www.postgresql.org/docs/current/mvcc-intro.html
- RocksDB snapshots: https://github.com/facebook/rocksdb/wiki/Snapshot
- FoundationDB snapshot read discussion: https://forums.foundationdb.org/t/does-snapshot-provide-snapshot-level-isolation/301
- Automerge binary format heads: https://automerge.org/automerge-binary-format-spec/
- Automerge synchronization protocol: https://posit-dev.github.io/automerge-r/articles/sync-protocol.html

## Why this matters to Ascend

Agents and UI clients need stable snapshots and explainable refresh boundaries when monitoring changing workbooks. MVCC and LSM systems model this with explicit versions, retained histories, and compaction watermarks. Ascend has pieces of that model, but they are split across compact range reads and interactive viewport sessions.

## Probe/implementation

- Inspected `packages/sdk/src/sheet-handle.ts`: compact reads keep one snapshot per compact window key and invalidate older tokens as `base-token-stale`.
- Inspected `packages/sdk/src/types.ts`: compact invalidations include `baseToken`, `changeToken`, reason, and required action, but not retained generation metadata.
- Inspected `packages/sdk/src/session.ts` references from the prior experiment: interactive viewport patches already have their own token and invalidation protocol.
- Inspected `apps/api/src/server.ts`, `apps/mcp/src/index.ts`, and TUI session/controller call sites to confirm compact reads and interactive sessions are separate public behaviors.
- Ran the existing compact token probe:

```bash
bun test packages/sdk/src/sdk.test.ts -t "compact changedSince"
```

## Results

The existing compact changedSince tests passed: 2 tests, 15 expectations, 0 failures.

The audit found that a narrow patch adding `currentGeneration` to compact `changeInvalidation` would be technically small, but incomplete:

- SDK compact reads use a latest-snapshot map, not a retained generation ledger.
- API and MCP merely forward compact invalidation objects.
- TUI mostly reads compact snapshots directly and does not consume the API/MCP compact invalidation shape.
- Interactive session viewport patches already have separate invalidation semantics and many tests.

## Confidence

High that this should not be folded as a one-surface tweak. Medium that the right next implementation is an internal shared `ViewportPatchLedger`, because the prior probe showed the model but not production integration cost.

## Fold-in decision

Promote to topic synthesis and future product/performance loop, but do not fold into production in this cycle. A partial generation field would look useful while leaving the harder retained-history behavior unresolved.

## Next question

Can a shared internal `ViewportPatchLedger` prototype cover both compact range reads and interactive session viewports with retained generation windows and explicit compaction reasons?
