# Fixture Policy Evidence Handoff

## Question

Can the product fixture gates for the top release claims carry current tracked-corpus scan evidence directly in the owner handoff?

## Hypothesis

Yes. The fixture scan scripts already produce machine-readable summaries. Folding compact summaries into `release-proof-index --owner-handoffs-json` should make the product decision concrete without silently accepting generated structural fixtures.

## External sources checked

- GitHub repository limits: https://docs.github.com/en/repositories/creating-and-managing-repositories/repository-limits
- GitHub large file guidance: https://docs.github.com/en/repositories/working-with-files/managing-large-files/about-large-files-on-github
- OpenSSF Scorecard checks: https://github.com/ossf/scorecard/blob/main/docs/checks.md
- SLSA provenance: https://slsa.dev/spec/v1.0-rc1/provenance
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

The top claims still depend on generated structural fixtures for signed and unknown package cases. Ascend needs to be honest about what the tracked public corpus proves, what it does not prove, and when product must explicitly accept generated evidence instead of treating it as a public binary replacement.

## Probe/implementation

Ran:

- `bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json`
- `bun run fixtures/benchmarks/package-action-fixture-scan.ts --json`

Folded scan summaries into `ReleaseProofIndexResult.fixturePolicyEvidence` and `releaseProofOwnerHandoffIndex.fixturePolicyEvidence`.

The handoff now reports:

- scan corpus
- scanned/rejected counts
- replacement status
- current generated structural cases
- safe-open signature/unknown match count
- package-action missing replacement features
- package-action feature counts
- `ownerApprovalRequired=true`

## Results

Current owner handoff probe:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json | jq '{headlineClaimsAllowed, implementationSurfacePromotionAllowed, missingRequirementCount, fixturePolicyEvidence}'
```

Observed:

- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- `fixturePolicyEvidence.status=tracked-scan-complete-owner-approval-required`
- `fixturePolicyEvidence.allScansUseTrackedCorpus=true`
- `fixturePolicyEvidence.publicReplacementGapsRemain=true`
- safe-open scanned 223 tracked fixtures, rejected 1, and found 0 signature/unknown replacements
- package-action scanned 223 tracked fixtures, rejected 1, and found `signaturePackage=0` and `syntheticUnknownPathFamily=0`

Targeted test:

- `bun test fixtures/benchmarks/release-proof-index.test.ts`

## Confidence

High that tracked-corpus scan evidence is now machine-visible in the owner handoff. Medium for the product decision because this does not prove no suitable public fixtures exist elsewhere and does not perform license/privacy review.

## Fold-in decision

Folded into release proof indexing as owner-decision evidence. Keep `public-edge-fixtures` and `edge-fixture-policy` missing until product either accepts disclosed generated structural fixtures or supplies approved public binary replacements.

## Next question

Should the next loop focus on release compact-report publication policy, or should it produce a product acceptance packet for generated structural fixtures using this scan evidence?
