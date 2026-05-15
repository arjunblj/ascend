# Release Fixture Policy Approval Checklist

## Question

Can the fixture-policy handoff include concrete product/release approval checklist items without marking any release gates satisfied?

## Hypothesis

Yes. The checklist can list pending owner decisions, validation commands, acceptance evidence, and rejection conditions while keeping `headlineClaimsAllowed=false` and all owner gates missing.

## External sources checked

- SPDX license list: https://spdx.org/licenses/
- OpenSSF Scorecard binary artifacts check: https://github.com/ossf/scorecard/blob/main/docs/checks.md
- GitHub repository limits: https://docs.github.com/en/repositories/creating-and-managing-repositories/repository-limits
- SLSA provenance: https://slsa.dev/spec/v1.0-rc1/provenance
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

Product and release owners need exact accept/reject conditions for generated structural fixtures. Without a checklist, research can keep producing useful evidence but still leave the claim board vague at the moment of owner review.

## Probe/implementation

Added `fixturePolicy.approvalChecklist` to `fixtures/benchmarks/release-proof-index.ts`. It contains four pending decisions:

1. product approval for safe-open generated signed/unknown structural fixtures;
2. product approval for package-action generated signature/unknown structural fixtures;
3. release approval for safe-open publication boundary wording;
4. release approval for package-action provenance boundary wording.

Each item carries `status: pending-owner-decision`, a validation command, acceptance evidence, and `rejectIf` text. The Markdown release proof index renders the checklist under `## Fixture Policy`.

## Results

`bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json` now includes `fixturePolicy.approvalChecklist` while the gate remains fail-closed:

- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- all checklist items are `pending-owner-decision`

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check --write fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
```

## Confidence

High that the checklist makes owner review more actionable without widening any surface. Medium that the exact wording is final; product/release owners may still change criteria or require stronger fixture acquisition.

## Fold-in decision

Promote to product/release proof packaging only. Do not mark `public-edge-fixtures`, `edge-fixture-policy`, `publication-boundary`, or `provenance-boundary` satisfied.

## Next question

Can the performance-owned `release-latency-run` and `streaming-matrix-boundary` blockers get the same approval-checklist treatment, including validation commands and rejection conditions, without promoting performance thresholds?
