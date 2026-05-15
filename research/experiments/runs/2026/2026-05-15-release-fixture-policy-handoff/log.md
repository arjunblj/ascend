# Release Fixture Policy Handoff

## Question

Can the top fixture blockers be converted from scattered prose into a machine-readable owner policy without marking the claims ready?

## Hypothesis

Yes. A compact fixture policy in the release proof index can make generated structural fixture acceptance explicit while keeping the release gate fail-closed.

## External sources checked

- GitHub repository limits: https://docs.github.com/en/repositories/creating-and-managing-repositories/repository-limits
- GitHub large-file guidance: https://docs.github.com/github/managing-large-files/working-with-large-files
- OpenSSF Scorecard checks, including binary artifacts: https://github.com/ossf/scorecard/blob/main/docs/checks.md
- SLSA provenance specification: https://slsa.dev/spec/v1.0-rc1/provenance
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

The top safe-open and package-action blockers are not implementation gaps. They are owner decisions about when generated structural package fixtures can support guarded release wording. Putting the policy into `release-proof-index` makes the decision reviewable by product/release owners and prevents research from silently promoting generated evidence.

## Probe/implementation

Added `fixturePolicy` to the release proof index and compact owner-handoff JSON. It records:

- current decision: `owner-approval-required`;
- generated structural fixture acceptance criteria;
- cases that require public binary fixtures before stronger claims;
- tracked scan commands for safe-open and package-action fixture replacement;
- current generated structural cases for each top artifact;
- external policy/provenance references;
- a boundary that says the policy is not publication approval, signed provenance, or license/privacy review.

## Results

`bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json` now includes a top-level `fixturePolicy` alongside `nextOwnerActions` and `implementationHandoffs`. The gate remains fail-closed:

- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- top owner actions remain `edge-fixture-policy` and `public-edge-fixtures`

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
```

## Confidence

High that this improves the owner handoff without changing production behavior. Medium that the criteria are complete; product/release owners may still refine thresholds for public binary acquisition, license review, or artifact publication.

## Fold-in decision

Promote to product/release proof packaging only. Do not mark the fixture gates satisfied and do not add new SDK, CLI, API, or MCP surfaces.

## Next question

Can the owner-policy handoff be paired with a compact approval checklist for product and release loops, while still keeping the release gate blocked until those owners explicitly approve it?
