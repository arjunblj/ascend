# Release Surface Promotion Fail-Closed

## Question

Can the release proof index make it machine-readable that unresolved proof blockers do not authorize new SDK, CLI, API, or MCP surfaces?

## Hypothesis

Yes. If the aggregate readiness summary carries an explicit `implementationSurfacePromotionAllowed` gate and boundary text, downstream owner loops can consume the proof index without inferring that missing evidence should be solved by adding another product surface.

## External sources checked

- SLSA provenance defines provenance as attestation evidence about how artifacts were produced, which is stronger than this local proof index: https://slsa.dev/spec/v0.1/provenance
- SLSA v1 provenance frames provenance as a build-platform attestation whose fields must be verified downstream: https://slsa.dev/provenance/v1-rc2
- The in-toto attestation framework documents a metadata format for supply-chain attestations and validation of origins: https://github.com/in-toto/attestation
- The in-toto project describes artifacts and rules for supply-chain integrity, which is outside the current Ascend release proof index boundary: https://github.com/in-toto/in-toto

## Why this matters to Ascend

The current top claims, "safe unknown workbook opening" and "auditable package-part mutation", are now blocked by owner decisions, validation runs, optional harness expansion, and release publication policy. Leaving that as prose creates a research failure mode: every missing proof becomes an excuse to add another surface. A fail-closed gate keeps Ascend's North Star focused on credible product claims backed by proof.

## Probe/implementation

Added two fields to `ReleaseProofReadinessSummary` in `fixtures/benchmarks/release-proof-index.ts`:

- `implementationSurfacePromotionAllowed`
- `implementationSurfacePromotionBoundary`

The gate currently follows `headlineClaimsAllowed`: when unresolved readiness requirements remain, implementation surface promotion is false. The markdown renderer now emits the gate and boundary in the release readiness section. Tests assert both JSON and Markdown output.

Commands run:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

## Results

- Focused test passed: 3 tests, 67 assertions.
- Release proof JSON now reports `implementationSurfacePromotionAllowed: false`.
- Boundary text explicitly states that current blockers do not authorize new SDK, CLI, API, or MCP surfaces.
- The top two proof artifacts remain blocked by 9 readiness requirements, so no new claim or product surface was promoted.

## Confidence

High for the gate semantics and harness behavior. Medium for owner-loop usefulness until product/correctness/performance/release owners consume the new field in their handoff workflows.

## Fold-in decision

Fold into the release proof harness. This is a tiny proof-owner guardrail, not a product surface. The handoff remains: safe-open to product/performance/release owners, package-action proof to correctness/product/performance/release owners.

## Next question

Can the top-two owner prompts consume `implementationSurfacePromotionAllowed=false` directly so future loops start from owner approval/validation instead of surface design?
