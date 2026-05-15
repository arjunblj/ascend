# Retained viewport patch contract

## Question

Can retained viewport patch history be packaged as an honest product claim with clear recovery guidance for invalid, stale, expired, and mismatched tokens?

## Hypothesis

Yes for compact reads and SDK interactive sessions, if the claim is bounded: Ascend retains recent viewport/window snapshots by key and returns either a patch or a fresh window/snapshot with an explicit invalidation reason and required action.

## External sources checked

- [PostgreSQL MVCC introduction](https://www.postgresql.org/docs/17/mvcc-intro.html) frames stable reader snapshots over changing data.
- [RocksDB iterator docs](https://github.com/facebook/rocksdb/wiki/Iterator) describe snapshot reads and retained files while iterators pin data, a useful contrast for bounded retention and expiration.
- [Yjs document updates](https://docs.yjs.dev/api/document-updates) expose incremental updates and state vectors, clarifying what Ascend is not claiming: collaborative CRDT convergence.
- [Automerge sync docs](https://automerge.org/automerge/automerge/sync/index.html) contrast full local-first history and sync with Ascend's narrower retained viewport history.

## Why this matters to Ascend

Agents and UI clients need to poll workbook windows without resending entire ranges after every edit. The claim is valuable only if recovery is deterministic: when a patch cannot be computed, clients must know whether to consume the returned full window, refresh the viewport, or discard stale tokens.

## Probe/implementation

- Inspected compact read surfaces in SDK, API, MCP, and docs.
- Inspected SDK interactive viewport patch tests and invalidation reasons.
- Ran focused tests for:
  - SDK compact `changedSince` retention and invalidation.
  - API compact `changedSince` invalidation.
  - MCP compact `changedSince` invalidation.
  - SDK interactive cross-session, expired, stale, invalid, and viewport-invalidated token handling.
- Ran a local generated-workbook probe comparing a one-cell compact patch against a full compact window.
- No production changes were made.

## Results

Focused validation passed:

```bash
bun test packages/sdk/src/sdk.test.ts -t "compact changedSince"
bun test apps/api/src/server.test.ts -t "compact changedSince"
bun test apps/mcp/src/index.test.ts -t "compact changedSince"
bun test packages/sdk/src/interactive-contract.test.ts -t "interactive viewport tokens from other sessions|interactive viewport patch results expose invalidation reasons|interactive pull patches reject tokens older"
```

Local compact-read probe:

| Case | Cells returned | JSON bytes | Recovery |
| --- | ---: | ---: | --- |
| Base full compact window | 100 | 11536 | Store returned `changeToken` |
| One-cell retained patch | 1 | 686 | Apply patch and store returned `changeToken` |
| Invalid token | full fresh window | n/a | `base-token-invalid`, `use-returned-window` |

Contract matrix:

| Surface | Patch when retained | Invalid token | Missing/cross-session token | Expired token | Metadata/layout invalidation |
| --- | --- | --- | --- | --- | --- |
| SDK compact read | yes | `base-token-invalid` | `base-snapshot-missing` | `base-token-expired` | request/window mismatch refresh |
| API compact read | yes | `base-token-invalid` | `base-snapshot-missing` | covered by SDK behavior | returned full window |
| MCP compact read | yes | `base-token-invalid` | `base-snapshot-missing` | covered by SDK behavior | returned full window |
| SDK interactive viewport | yes | `base-token-invalid` | `base-snapshot-missing` | `base-token-expired` | `viewport-invalidated` |

## Confidence

High for the bounded SDK/API/MCP compact-read claim and SDK interactive claim. Medium for a broader product claim because there is no CLI `changedSince` surface and no telemetry showing default retention depth is ideal.

## Fold-in decision

Promote to topic synthesis and product/DX docs. Do not add retention knobs yet. The next product loop should publish recovery examples and keep the boundary explicit: this is bounded patch history, not collaboration or unbounded event sourcing.

## Next question

Should formula language-service primitives be promoted next, or should the loop stay on claim packaging and create the product-ready proof ladder for token-bounded agent view plus retained patch history?
