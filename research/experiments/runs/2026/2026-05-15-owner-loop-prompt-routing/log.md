# Owner Loop Prompt Routing

## Question

Can Ascend separate remaining release blockers into owner-specific next-loop prompts, so research stops handing broad themes to implementation?

## Hypothesis

Yes. The current owner-handoff JSON already classifies every missing release gate by owner loop, priority, and next-step kind. A local grouping probe should produce clear product, correctness, performance, and release prompts without changing production code.

## External sources checked

- GitHub artifact attestation documentation: https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
- GitHub artifact attestation concepts: https://docs.github.com/en/enterprise-cloud@latest/actions/concepts/security/artifact-attestations
- SLSA provenance model: https://slsa.dev/spec/v1.0-rc1/provenance
- Microsoft Protected View documentation: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- OpenSSF Scorecard binary artifacts check: https://github.com/ossf/scorecard/blob/main/docs/checks.md

## Why this matters to Ascend

The North Star requires proof-backed release claims. The top risk in the research loop is continuing to add surfaces while owner gates are still blocked. Owner-specific prompts make the next work concrete: product decides fixture evidence, correctness approves unsupported-feature boundaries, performance runs the approved validation, and release defines publication policy.

## Probe/implementation

Local probe:

```bash
bun --eval 'import { runReleaseProofIndex, releaseProofOwnerHandoffIndex } from "./fixtures/benchmarks/release-proof-index.ts"; const handoff=releaseProofOwnerHandoffIndex(await runReleaseProofIndex({includeTimings:false})); const grouped={}; for (const action of handoff.nextOwnerActions) { (grouped[action.ownerLoop] ??= []).push({artifact: action.artifact, requirementId: action.requirementId, nextStepKind: action.nextStepKind, priority: action.priority, acceptanceEvidence: action.acceptanceEvidence, forbiddenShortcut: action.forbiddenShortcut}); } console.log(JSON.stringify({releaseGate: handoff.releaseGate, headlineClaimsAllowed: handoff.headlineClaimsAllowed, implementationSurfacePromotionAllowed: handoff.implementationSurfacePromotionAllowed, missingRequirementCount: handoff.missingRequirementCount, grouped}, null, 2));'
```

Implementation:

- Updated `research/experiments/syntheses/2026-05-owner-handoff.md` with owner-loop prompts only.
- No production code changes.

## Results

The grouped probe reported:

- `releaseGate=blocked-by-publication-policy`
- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- product blockers: `package-action-proof/edge-fixture-policy`, `safe-open-proof/public-edge-fixtures`
- correctness blocker: `package-action-proof/unsupported-feature-boundary`
- performance blockers: `safe-open-proof/release-latency-run`, `package-action-proof/streaming-matrix-boundary`
- release blockers: package-action provenance boundary, safe-open publication boundary, and both compact-report publication policies

The synthesis now contains four next-loop prompts and keeps all lower-ranked research directions out of implementation until these owner gates move.

## Confidence

High. The prompt routing is a direct projection of machine-readable owner-handoff JSON. It does not create a new source of truth.

## Fold-in decision

Promote to topic synthesis only. The owner-loop prompt board should guide the next product/correctness/performance/release loops, but it does not satisfy any gate.

## Next question

If the release/packageability RC gate work in the worktree becomes coherent, validate it separately as a release-loop candidate; otherwise continue with research-only claim stewardship until the dirty worktree clears.
