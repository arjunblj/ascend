# Release Owner Handoff Data

## Question

Can the release proof index hand off only the top two implementation-loop claims as machine-readable proof work, while still failing closed on new surface promotion?

## Hypothesis

Yes. The proof index already knows the top artifacts, proof commands, blocking requirements, and owner action kinds. Emitting an `implementationHandoffs` array from the same canonical result should reduce drift between synthesis docs and implementation loops.

## External sources checked

- SLSA provenance notes that provenance fields must be verified downstream, which supports keeping handoff gates explicit rather than implicit: https://slsa.dev/spec/v1.0-rc1/provenance
- SLSA source requirements describe published verification data as evidence for downstream checks and policies: https://slsa.dev/spec/v1.2/source-requirements
- The in-toto attestation spec describes authenticated metadata over subjects and predicates; Ascend's local handoff data remains below that bar: https://github.com/in-toto/attestation/blob/main/spec/README.md
- in-toto Archivista stores attestations and indexes subjects for retrieval; Ascend is only emitting local proof-handoff metadata, not a storage or attestation service: https://github.com/in-toto/archivista

## Why this matters to Ascend

The claim stewardship rule says hand off only the top one or two product claims. Without a machine-readable handoff, a later loop can accidentally reopen lower-ranked surfaces or treat prose as permission to add APIs. A proof-index handoff keeps the next loops pointed at owner decisions, validation, and boundary approval.

## Probe/implementation

Added `implementationHandoffs` to `ReleaseProofReadinessSummary` in `fixtures/benchmarks/release-proof-index.ts`. Each handoff includes:

- rank;
- artifact and product-shaped claim;
- owner loops;
- proof and compact-report commands;
- `implementationSurfacePromotionAllowed`;
- blocking `readyWhen` requirement IDs;
- next-step kinds;
- boundary text that says the handoff is not permission to add SDK, CLI, API, or MCP surfaces.

The Markdown renderer now emits a compact handoff line, and the focused test asserts the top two handoffs.

Commands run:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

## Results

- Focused test passed: 3 tests, 71 assertions.
- The proof index now emits exactly two implementation handoffs:
  - rank 1: `safe-open-proof`, claim `safe unknown workbook opening`;
  - rank 2: `package-action-proof`, claim `auditable package-part mutation`.
- Both handoffs carry `implementationSurfacePromotionAllowed=false`.
- The handoffs name owner loops and blocker IDs, so future loops can start at approval/validation instead of new surface design.

## Confidence

High for the harness output and ordering. Medium for downstream adoption until owner loops consume this data directly.

## Fold-in decision

Fold into the release proof harness. This is a proof-routing improvement only; it does not promote any product surface.

## Next question

Can the owner loops resolve the first blocker class, product acceptance of disclosed generated fixtures versus public binary replacement, without changing implementation surfaces?
