# Release Trust Owner Handoff Boundary

## Question

Should the release proof owner handoff expose the correctness/trust completeness boundary, rather than leaving it only in `docs/RELEASE_TRUST_MATRIX.md`?

## Hypothesis

If owner handoff JSON includes the out-of-scope correctness boundary, future loops can see that broad formula breadth, UX orchestration, performance tuning, malformed-field enumeration, and unknown-feature implementation do not change A+ release status unless they prove a concrete trust failure path. This should reduce accidental scope expansion without adding a new product surface.

## External sources checked

- SLSA distributing provenance: https://slsa.dev/spec/v1.0/distributing-provenance
- GitHub artifact attestations: https://docs.github.com/en/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds
- Microsoft Protected View: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653

## Why this matters to Ascend

Ascend's top release claims are blocked by explicit owner gates, not by lack of more feature surfaces. Encoding the trust-completeness boundary in owner handoff JSON makes the release proof index a stronger routing artifact: correctness owners can reject work that does not affect silent corruption, exact journals, or post-write drift.

## Probe/implementation

Added `trustCompletenessBoundaryEvidence` to `fixtures/benchmarks/release-proof-index.ts` and surfaced it in both full Markdown and `--owner-handoffs-json`.

The new evidence records:

- validation command: `bun test packages/sdk/src/release-trust-matrix.test.ts`
- matrix path: `docs/RELEASE_TRUST_MATRIX.md`
- five out-of-scope classes and promotion criteria
- gates it does not close: product, performance, release
- source references for provenance and Protected View boundaries

No SDK, CLI, API, MCP, workbook mutation, formula rename, viewport, or benchmark surface was added.

## Results

The owner handoff now emits a machine-readable `trustCompletenessBoundaryEvidence` object with `status: boundary-pinned-owner-scope`. Markdown output includes a "Correctness/trust completeness boundary" section. Release gate behavior stays unchanged: this is owner routing evidence only and does not satisfy publication gates.

Validation:

- `bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`
- `bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`
- `bunx tsc --build`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`

## Confidence

High that this changes owner routing without expanding release scope. Medium that the boundary list is complete; the list is intentionally pinned to the current release trust matrix and can be revised only with proof of release-trust impact.

## Fold-in decision

Fold into release-proof-index owner handoff. This directly changes gate/owner-task visibility and remains below product-surface promotion.

## Next question

Should `nextOwnerActions` explicitly include a correctness action for accepting this boundary, or is the current `trustCompletenessBoundaryEvidence` sufficient because it is non-blocking scope control?
