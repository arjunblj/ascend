# Release Owner Action Kind Gate

## Question

Can the release proof index distinguish owner decisions from implementation work so research does not keep promoting new surfaces when the top blockers are policy or validation gates?

## Hypothesis

Yes. The release proof index already ranks missing owner actions. Adding a machine-readable `nextStepKind` to each owner action should make the top handoff explicit: fixture approval/replacement is not the same as adding another SDK/CLI/API/MCP surface.

## External sources checked

- Open Packaging Conventions define package parts, relationships, and digital signature boundaries, supporting Ascend's per-part proof model while keeping signer/origin validation outside local proof: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft Protected View frames unsafe opening as a trust/protection UX; Ascend's safe-open claim must remain package-feature routing, not malware scanning: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- SheetJS CE write docs describe writer options and data preservation scope, useful as competitor contrast for why Ascend's stronger claim is explicit package accounting: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

The release claim board says the top two claims are safe unknown workbook opening and auditable package-part mutation. The current blockers are mostly owner approval, fixture disclosure/replacement, validation runs, and publication policy. If the release index does not make that machine-readable, future research loops can mistake an owner decision for permission to build another surface.

## Probe/implementation

Folded a small proof-harness change into `fixtures/benchmarks/release-proof-index.ts`:

- `ReleaseProofNextOwnerAction` now includes `nextStepKind`.
- Fixture gates are tagged `owner-decision-or-fixture-replacement`.
- Unsupported-feature boundaries are tagged `owner-boundary-approval`.
- Release latency is tagged `validation-run`.
- Streaming matrix is tagged `owner-decision-or-harness-expansion`.
- Publication/provenance/compact-report gates are tagged `publication-policy`.
- Markdown `Next owner actions` now includes the action kind.

No SDK, CLI, API, MCP, or workbook mutation surface was added.

## Results

Focused validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
```

Observed proof output:

- `releaseGate`: `blocked-by-publication-policy`
- `headlineClaimsAllowed`: `false`
- top package-action gate: `owner-decision-or-fixture-replacement`
- top safe-open gate: `owner-decision-or-fixture-replacement`
- release latency gate: `validation-run`
- streaming matrix gate: `owner-decision-or-harness-expansion`
- release/provenance/compact-report gates: `publication-policy`

## Confidence

High for the release-index proof shape. Medium for owner-loop resolution because this change clarifies the handoff; it does not itself approve generated fixtures, publish compact reports, or create release-environment latency evidence.

## Fold-in decision

Folded into the release proof benchmark harness. This is a proof-gate improvement, not a product surface. Keep top implementation handoffs unchanged.

## Next question

Can the top product/performance and correctness/product loops now consume the release index owner-action kinds directly instead of asking research to create another broad synthesis?
