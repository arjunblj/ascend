# Generated Fixture Owner Action Proof

## Question

Can the release proof index give product owners a concrete recommended action for each generated safe-open and package-action fixture without marking the fixture gates satisfied or promoting new product surfaces?

## Hypothesis

Yes. The index already records generated fixture provenance, replacement scan evidence, allowed use, and forbidden use. Adding a per-case recommended owner action should make the handoff more actionable while preserving `headlineClaimsAllowed=false` and `implementationSurfacePromotionAllowed=false`.

## External sources checked

- GitHub repository limits recommend keeping generated files outside Git or in LFS when appropriate: https://docs.github.com/en/repositories/creating-and-managing-repositories/repository-limits
- GitHub large file storage documents binary-file handling outside normal Git object storage: https://docs.github.com/github/managing-large-files/about-git-large-file-storage
- OpenSSF Scorecard documents binary-artifact risk in source repositories: https://github.com/ossf/scorecard/blob/main/docs/checks.md
- SLSA provenance defines provenance as attested artifact/build metadata, which Ascend's local proof index does not produce: https://slsa.dev/spec/v1.0-rc1/provenance
- GitHub artifact attestations describe build provenance attestations; Ascend's generated fixture proof is not an attestation: https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations

## Why this matters to Ascend

The top two release claims are blocked mainly by owner fixture decisions. Without a recommended action, future loops may keep searching broadly for public signed/unknown workbooks or may accidentally treat generated packages as accepted public evidence. The useful proof is to make the owner decision explicit and fail-closed.

## Probe/implementation

Inspected `fixtures/benchmarks/release-proof-index.ts` and its tests. Added `recommendedOwnerAction` to each `generatedFixtureDecisionEvidence.cases[]` row:

- safe-open `signed`
- safe-open `unknown-part`
- safe-open `malformed`
- package-action `signature-invalidation-drop`
- package-action `unknown-part-error`

The recommendation wording accepts only disclosed local topology/rejection proof and keeps trust, provenance, malware, recovery, and arbitrary preservation wording forbidden.

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bunx tsc --build
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
```

## Results

The owner-handoff JSON now emits `recommendedOwnerAction` for all five generated structural fixture cases. The rerun still reports:

- `releaseGate=blocked-by-publication-policy`
- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- `generatedFixtureDecisionEvidence.ownerApprovalRequired=true`
- `generatedFixtureDecisionEvidence.publicReplacementGapsRemain=true`

This is a proof-routing improvement only. It does not satisfy `public-edge-fixtures`, `edge-fixture-policy`, publication policy, provenance policy, or release latency policy.

## Confidence

High for the scoped fold-in: tests cover the new field, the handoff JSON includes it, and the release gates remain fail-closed. Medium for owner acceptance because product/release owners still need to decide whether disclosed generated topology proof is good enough for guarded local release proof wording.

## Fold-in decision

Fold into the benchmark/release proof harness as owner-routing evidence. Do not promote production surfaces. The owner action recommendation should guide product/release approval, not replace it.

## Next question

Can the performance owner reduce the `safe-open-proof/release-latency-run` blocker with a tracked-clean timed proof packet that reports public-input timing without turning it into a release threshold?
