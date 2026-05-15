# Viewport Patch Proof Harness

## Question

Can retained viewport patch history get a deterministic proof harness for valid token, skipped token, expired token, projection change, cross-session token, and metadata invalidation without implying collaboration?

## Hypothesis

Yes. `AscendSession` already exposes viewport `changeToken`, retained patch results, and explicit invalidation reasons. A small harness can package the product claim as bounded patch retention plus required refresh actions, while avoiding CRDT, multi-writer, or unbounded-history language.

## External sources checked

- PostgreSQL MVCC docs describe snapshot-based reads over older database versions, supporting the idea of retained read versions with visibility boundaries: https://www.postgresql.org/docs/17/mvcc-intro.html
- PostgreSQL snapshot docs describe `xmin`, `xmax`, and active transaction lists as explicit snapshot metadata: https://www.postgresql.org/docs/9.5/functions-info.html
- Automerge docs distinguish local-first CRDT collaboration from ordinary patch/snapshot plumbing: https://automerge.org/docs/reference/concepts/
- Yjs docs describe document updates and state vectors for collaborative synchronization, which Ascend should not imply from retained viewport patches: https://docs.yjs.dev/api/document-updates

## Why this matters to Ascend

The release claim board says retained viewport patch history is credible but needs proof packaging. Spreadsheet clients and agents can miss responses, retry polls, or change viewport projections. The useful product claim is narrow: Ascend can patch from bounded retained tokens and return machine-readable reasons when it cannot. That is not collaboration or unlimited event sourcing.

## Probe/implementation

- Inspected `packages/sdk/src/session.ts`, `packages/sdk/src/interactive-contract.test.ts`, retained compact-read logs, and the release claim board.
- Added `fixtures/benchmarks/viewport-patch-proof.ts`.
  - Runs SDK interactive sessions against a generated workbook.
  - Emits Markdown by default and JSON with `--json`.
  - Covers retained patch, skipped older token, invalid token, cross-session token, expired history, projection change, and metadata invalidation.
  - Records changed refs, patch bytes, invalidation reasons, and required action.
- Added `fixtures/benchmarks/viewport-patch-proof.test.ts`.
  - Asserts retained patch behavior.
  - Asserts explicit invalidation reasons.
  - Asserts report wording rejects CRDT/collaboration claims.

## Results

Local proof command:

```bash
bun run fixtures/benchmarks/viewport-patch-proof.ts
```

| Case | Expected | Observed | Passed | Patch bytes | Changed refs | Invalidation reason | Required action |
| --- | --- | --- | --- | ---: | --- | --- | --- |
| retained-patch | patch:A1 | patch:A1 | true | 315 | A1 | n/a | n/a |
| skipped-token-retained | patch from older retained token | patch:A1 | true | 315 | A1=3 | n/a | n/a |
| invalid-token | base-token-invalid | base-token-invalid | true | 0 | n/a | base-token-invalid | use-returned-snapshot |
| cross-session-token | base-snapshot-missing | base-snapshot-missing | true | 0 | n/a | base-snapshot-missing | use-returned-snapshot |
| expired-history | base-token-expired | base-token-expired | true | 0 | n/a | base-token-expired | use-returned-snapshot |
| projection-change | base-snapshot-missing | base-snapshot-missing | true | 0 | n/a | base-snapshot-missing | use-returned-snapshot |
| metadata-invalidation | viewport-invalidated | viewport-invalidated | true | 0 | n/a | viewport-invalidated | use-returned-snapshot |

Boundary:

- Projection changes currently surface as `base-snapshot-missing`, not a distinct projection-specific reason. That is acceptable for recovery, but product wording should say "projection change returns a fresh snapshot with an invalidation reason," not promise a dedicated projection code.
- Cross-session tokens are not shared history. They return a fresh snapshot path.

Validation passed:

- `bun test fixtures/benchmarks/viewport-patch-proof.test.ts`
- `bunx biome check fixtures/benchmarks/viewport-patch-proof.ts fixtures/benchmarks/viewport-patch-proof.test.ts`
- `bunx tsc --build`
- `bun run test:changed` (5007 pass, 1 skip)

## Confidence

High for the SDK interactive viewport claim. Medium for cross-surface product packaging because compact SDK/API/MCP reads and interactive SDK viewports have related but not identical patch protocols.

## Fold-in decision

Fold into product/performance proof as a tracked harness. Do not add retention knobs or collaboration language. The next product loop should package examples over existing SDK/API/MCP compact reads and SDK interactive sessions.

## Next question

Should formula language-service primitives get a proof harness that demonstrates diagnostics, hover, completions, reference spans, binding roles, and prepare-rename refusals without implementing edit-producing rename?
