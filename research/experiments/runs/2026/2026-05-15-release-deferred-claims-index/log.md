# Release Deferred Claims Index

## Question

Can the release proof index carry a machine-readable "do not promote yet" list for lower-ranked research directions?

## Hypothesis

Yes. If the same canonical index that hands off the top two claims also lists deferred claims, their owner loops, proof gaps, kill criteria, and boundaries, future loops are less likely to convert promising research into product surfaces before proof exists.

## External sources checked

- LSP 3.17 separates `prepareRename` from edit-producing rename and allows refusal semantics, supporting the formula rename freeze: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Apache Arrow documents a language-independent columnar format, supporting the sidecar research direction while keeping it below product promotion: https://arrow.apache.org/docs/format/Columnar.html
- PostgreSQL MVCC documentation frames snapshots and multiversion consistency, supporting viewport history analogies while preventing collaboration/transaction-isolation overclaims: https://www.postgresql.org/docs/17/mvcc-intro.html
- The Model Context Protocol repository points to the official MCP documentation and standardizes tools/resources, supporting agent workflow observability as an agent-DX research direction: https://github.com/modelcontextprotocol/modelcontextprotocol

## Why this matters to Ascend

The current owner handoff should stay focused on "safe unknown workbook opening" and "auditable package-part mutation." Formula intelligence, token-bounded views, retained patch history, sidecars, oracle routing, and agent traces are valuable but should not compete with the top proof loops. A machine-readable deferred list makes that portfolio decision explicit.

## Probe/implementation

Added `deferredClaims` to `ReleaseProofIndexResult` in `fixtures/benchmarks/release-proof-index.ts`.

The index now lists six non-top claims:

- formula language-service primitives;
- token-bounded agent view;
- retained viewport patch history;
- columnar scan sidecars;
- formula oracle routing;
- agent workflow observability.

Each row includes status, owner loops, reason, proof needed, kill criterion, and boundary. The Markdown renderer now emits a "Deferred Claims" table. The harness clones deferred rows before returning them so matcher objects in tests cannot mutate the shared constant.

Commands run:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
```

## Results

- Focused test passed: 3 tests, 82 assertions.
- Markdown output now includes a `Deferred Claims` table.
- JSON output reports `deferredClaimCount: 6`.
- The formula row explicitly freezes edit-producing rename.
- The columnar sidecar row stays `do-not-promote-yet`.

## Confidence

High for the proof-index behavior. Medium for the exact deferred list until owners decide whether agent-view or viewport patch history should become a release proof artifact later.

## Fold-in decision

Fold into the release proof harness. This is claim stewardship, not a production surface.

## Next question

Can owner loops consume the top-two handoffs and deferred claims from the JSON index directly, without relying on long Markdown synthesis documents?
