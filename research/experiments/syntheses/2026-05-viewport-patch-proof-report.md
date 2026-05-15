# 2026-05 Viewport Patch Proof Report

Date: 2026-05-15

## Claim

Ascend can patch viewport reads from bounded retained tokens and return explicit invalidation reasons when a patch cannot be produced.

## Claim Wording That Is Safe Today

Ascend has local proof for bounded single-session viewport patch retention in the SDK interactive session and for `changedSince` invalidation recovery in API and MCP compact reads. This is not collaboration, CRDT convergence, multi-writer sync, or unbounded event sourcing.

## External Contrast

- [PostgreSQL MVCC](https://www.postgresql.org/docs/17/mvcc-intro.html) frames consistent reads as snapshot-based visibility over retained versions.
- [SQLite WAL isolation](https://www.sqlite.org/isolation.html) documents snapshot isolation where a reader sees a stable database image even while writers append changes.
- [RocksDB snapshots](https://github.com/facebook/rocksdb/wiki/Snapshot) capture a point-in-time view tied to an internal sequence number and do not persist across restarts.
- [Automerge](https://automerge.org/docs/reference/concepts/) and [Yjs updates](https://docs.yjs.dev/api/document-updates) are collaboration/CRDT systems; Ascend's viewport patch claim should not borrow that language.

## Proof Status

| Required proof | Current evidence | Status |
| --- | --- | --- |
| SDK retained patch | `fixtures/benchmarks/viewport-patch-proof.ts` covers retained patch, skipped older token, invalid token, cross-session token, expired history, projection change, and metadata invalidation | Covered |
| API compact read | `apps/api/src/server.test.ts` covers changed window invalidation and changed source invalidation for `/read` compact `changedSince` | Covered for recovery, not retained interactive patch stream |
| MCP compact read | `apps/mcp/src/index.test.ts` covers changed window, selected-column projection, and changed source invalidation for `ascend.read` compact `changedSince` | Covered for recovery, not retained interactive patch stream |
| CLI surface | No current retained-patch or `changedSince` CLI surface was found | Missing; do not release-headline cross-surface claim yet |
| Invalidation vocabulary | `base-token-invalid`, `base-token-expired`, `base-snapshot-missing`, and `viewport-invalidated` are validated | Covered |
| Honest boundary | Report and harness reject CRDT, multi-writer, and unbounded-history language | Covered |

## Fresh Local Probe

Probe command:

```bash
bun run fixtures/benchmarks/viewport-patch-proof.ts
```

| Case | Expected | Observed | Passed | Patch bytes | Boundary |
| --- | --- | --- | --- | ---: | --- |
| retained-patch | `patch:A1` | `patch:A1` | true | 315 | SDK interactive retained token |
| skipped-token-retained | older retained token patches to latest visible cell | `patch:A1` | true | 315 | SDK keeps bounded retained history |
| invalid-token | `base-token-invalid` | `base-token-invalid` | true | 0 | Caller must use returned snapshot |
| cross-session-token | `base-snapshot-missing` | `base-snapshot-missing` | true | 0 | Tokens are not shared history |
| expired-history | `base-token-expired` | `base-token-expired` | true | 0 | Retention is bounded |
| projection-change | `base-snapshot-missing` | `base-snapshot-missing` | true | 0 | No dedicated projection reason yet |
| metadata-invalidation | `viewport-invalidated` | `viewport-invalidated` | true | 0 | Metadata edits force refresh |

Cross-surface validation:

```bash
bun test fixtures/benchmarks/viewport-patch-proof.test.ts
bun test apps/api/src/server.test.ts -t "compact changedSince"
bun test apps/mcp/src/index.test.ts -t "compact changedSince"
```

## Interpretation

- The SDK interactive claim is strong: retained tokens can produce compact patches, and invalid tokens or invalidated views return explicit refresh instructions.
- API and MCP compact reads support `changedSince` recovery with fresh windows and explicit invalidation metadata, but they are not the same retained interactive patch stream.
- CLI is the missing surface. Do not claim SDK/CLI/API/MCP parity unless a future product loop intentionally adds or rejects a CLI contract.
- Projection changes currently use `base-snapshot-missing`, not a dedicated projection-specific reason.

## Fold-In Recommendation

Promote to product/performance claim packaging only as a guarded claim: "SDK retained viewport patches, plus API/MCP compact read invalidation recovery." Do not promote to a release headline until the owner loop decides whether CLI needs a surface or the claim wording should exclude CLI.
