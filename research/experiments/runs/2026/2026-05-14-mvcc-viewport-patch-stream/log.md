# MVCC Viewport Patch Stream

## Question

Can UI patch streams use database MVCC/LSM ideas for generation-aware snapshots?

## Hypothesis

Yes. Ascend already has interactive viewport change tokens and patch invalidation, but it stores only the latest snapshot per viewport key. A retained, generation-indexed delta ledger could let clients ask for patches from older base tokens, compact old segments explicitly, and explain when a full refresh is required.

## External sources checked

- [PostgreSQL MVCC introduction](https://www.postgresql.org/docs/current/mvcc-intro.html): MVCC gives readers a stable snapshot while writers continue, avoiding read/write blocking.
- [RocksDB snapshots](https://github.com/facebook/rocksdb/wiki/Snapshot): snapshots are tied to sequence numbers; keys newer than a snapshot are invisible, and older/equal keys are visible.
- [RFC 6902 JSON Patch](https://www.rfc-editor.org/rfc/rfc6902.html): standardizes ordered partial document updates and explicit operation semantics.
- [Yjs docs](https://docs.yjs.dev/): local-first documents sync changes automatically and do not depend on update arrival order.
- [Automerge](https://automerge.org/): emphasizes durable change history, branching/merging, compact storage, and instantaneous local edits.

## Why this matters to Ascend

Large spreadsheet UIs need stable viewport reads while edits, recalculation, formatting, sorting, and hydration are happening. Agents also need compact deltas, not full viewport snapshots, when monitoring a workbook. A generation-aware patch stream can make Ascend feel like a database-backed UI system: clients read from a named snapshot, receive deltas until retention expires, and get a concrete reason when they must refresh.

This directly supports SOTA agent and human workflows: responsive TUI/web clients, watch-mode agents, and explainable UI synchronization after edits.

## Probe/implementation

Inspected local implementation:

- `packages/sdk/src/session.ts` defines `InteractiveViewportRequest`, `InteractiveViewportResult`, `InteractiveViewportPatch`, `changeToken`, `changedSince`, and invalidation reasons.
- `AscendSession.readViewport` currently stores one snapshot per viewport key in `viewportSnapshots`. If `changedSince` is not the latest token for that key, it returns `base-token-stale` even if the intermediate changes are still in `recentChanges`.
- `AscendSession.readViewportPatchResult` can produce a sparse patch from recorded refs when a mutable workbook is available, but still requires the exact retained snapshot token.
- `recentChanges` stores bounded generation -> ref sets, but it does not store per-viewport value/tombstone segments.
- `rebaseViewportSnapshots` rereads stored viewport snapshots after mutable workbook preparation and may invalidate with broad `refs: null`.
- `apps/api/src/server.ts` exposes compact `changedSince` reads, and `apps/api/src/server.test.ts` covers unchanged windows, invalid tokens, and source changes.

Added ignored probe `research/experiments/runs/2026/2026-05-14-mvcc-viewport-patch-stream/probes/mvcc-viewport-ledger.ts`. It models:

- viewport tokens with generation suffixes;
- append-only patch segments containing puts and tombstones;
- `patchSince(baseToken)` that merges segments from base generation to current generation;
- LSM-style compaction by advancing a retention watermark;
- explicit `base-token-expired` once compacted history is no longer available.

Validation command:

```bash
bun run research/experiments/runs/2026/2026-05-14-mvcc-viewport-patch-stream/probes/mvcc-viewport-ledger.ts
```

## Results

The probe produced three generations of viewport changes:

- `A1: 1 -> 10 -> 11`
- `B1` removed
- `D1` added

`patchSince(token0)` returned the final visible state for changed refs only:

| Changed ref | Value |
| --- | --- |
| `A1` | `11` |
| `D1` | `true` |

It also returned `removedRefs: ["B1"]`.

`patchSince(currentToken)` returned an empty patch. After compaction advanced retention past generation 1, `patchSince(token0)` returned `base-token-expired`.

The probe suggests a stronger protocol than "latest snapshot only": keep per-viewport delta segments for a bounded window, merge updates by ref, emit tombstones, and compact using an explicit watermark. That gives clients deterministic behavior:

- patch if the base generation is retained;
- refresh if the base token is invalid, from the future, or compacted away;
- attach generation and retention metadata to every patch.

## Confidence

Medium. The model is small but maps cleanly to Ascend's existing token and invalidation concepts. Real confidence needs a production benchmark with TUI/API viewport traces, recalculation changes, structural edits, sort/filter, and sheet metadata invalidation.

## Fold-in decision

Promote to product/DX loop and performance loop.

Recommended fold-in:

1. Add an internal `ViewportPatchLedger` behind `AscendSession`.
2. Store bounded patch segments per viewport key: `generation`, changed cell payloads, removed refs, and metadata-invalidating flags.
3. Preserve current full-snapshot response behavior as the fallback.
4. Let `changedSince` patch from any retained generation, not only the immediately previous token.
5. Report `retainedFromGeneration`, `baseGeneration`, `currentGeneration`, and compaction reason in invalidations.
6. Validate with API tests for stale-but-retained tokens, compacted tokens, structural viewport changes, and recalc-driven changed refs.

Do not fold into production here. The current implementation is already correct; this is a protocol/efficiency improvement.

## Next question

Can we build a release proof bundle that demonstrates inspect, plan, commit, reopen, diff, and audit on real files without fake claims?
