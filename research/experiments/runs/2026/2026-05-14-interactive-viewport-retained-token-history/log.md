# Interactive Viewport Retained Token History

## Question

Can interactive session viewport patches use a retained token ledger so skipped patch/read responses still patch while retained, without weakening existing invalidation guarantees?

## Hypothesis

Yes. `AscendSession` already keeps recent change refs and explicit invalidation reasons. Replacing its single latest viewport snapshot with a bounded per-viewport ledger should let retained base tokens produce patches while preserving missing, invalid, expired, and viewport-invalidated recovery paths.

## External sources checked

- PostgreSQL transaction snapshots: https://www.postgresql.org/docs/18/sql-set-transaction.html
- PostgreSQL concurrency control / MVCC chapter: https://www.postgresql.org/docs/18/mvcc.html
- RocksDB snapshots: https://github.com/facebook/rocksdb/wiki/Snapshot
- Automerge synchronization protocol heads: https://posit-dev.github.io/automerge-r/articles/sync-protocol.html
- Local-first software paper page: https://martin.kleppmann.com/2019/10/23/local-first-at-onward.html

## Why this matters to Ascend

Interactive spreadsheet clients and agents can miss an incremental response because of retries, transport gaps, or UI scheduling. A single-snapshot protocol forces a full viewport refresh for any skipped token even when Ascend still has enough local history to explain the delta. Retained viewport snapshots move Ascend closer to database and local-first systems: patch while the base is retained, explicitly refresh when it is not.

## Probe/implementation

- Inspected `packages/sdk/src/session.ts` around `AscendSession.readViewport()`, `readViewportPatchResult()`, `recentChanges`, `changedRefsSince()`, and `rebaseViewportSnapshots()`.
- Confirmed the current code stored one `viewportSnapshots` entry per viewport key and returned `base-token-stale` whenever the requested token was not the latest token.
- Folded in a scoped SDK implementation:
  - changed viewport snapshot storage to a bounded ledger retaining the last 8 snapshots per viewport key;
  - let `readViewport()` and `readViewportPatchResult()` locate any retained base token before diffing;
  - preserved `base-snapshot-missing`, `base-token-invalid`, `base-token-stale`, `base-token-expired`, and `viewport-invalidated` semantics;
  - classified tokens older than the retained ledger as `base-token-expired`;
  - rebased all retained snapshots during mutable-workbook preparation checks.
- Added contract coverage for retained base-token patching and ledger expiration.

## Results

Focused validation passed:

```bash
bun test packages/sdk/src/interactive-contract.test.ts -t "interactive viewport"
bun test packages/sdk/src/interactive-contract.test.ts
bunx biome check packages/sdk/src/session.ts packages/sdk/src/interactive-contract.test.ts research/experiments/index.md research/experiments/runs/2026/2026-05-14-interactive-viewport-retained-token-history/log.md
bunx tsc --build
bun run test:changed
```

The targeted tests now cover:

- cross-session tokens still require a full snapshot until that session has its own snapshot;
- retained base tokens can produce a patch after newer viewport tokens have been issued;
- old tokens that fall out of the retained viewport ledger return `base-token-expired`;
- invalid tokens, missing snapshots, recent-change expiration, and metadata invalidation still report explicit recovery reasons.

`bun run test:changed` first failed inside the sandbox because API tests could not bind local `Bun.serve()` ports. Rerunning the same gate outside the sandbox passed with 4974 tests, 1 skip, and 0 failures.

## Confidence

Medium. The change is scoped to SDK interactive session token bookkeeping and keeps the public result shape unchanged. The main residual risk is memory growth for many distinct viewport keys, bounded per key but not globally capped.

## Fold-in decision

Folded into the product/DX and performance loops. This is a production step toward generation-aware viewport patch streams and should be followed by telemetry or a global session snapshot budget before increasing retention.

## Next question

Can Ascend expose compact telemetry for retained snapshot ledgers so agents can tell when patch streams are healthy versus frequently refreshing?
