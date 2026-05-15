# Release Readiness Fail-Closed Summary

## Question

Can the top claim owner loops consume release proof readiness from one aggregate field and fail closed whenever any `readyWhen` requirement is still missing?

## Hypothesis

Yes. Per-artifact `readyWhen` checklists are useful for humans, but owner loops need a single machine-readable gate. The release proof index can summarize missing requirements by owner loop and artifact, with `headlineClaimsAllowed: false` until every readiness requirement is satisfied and every artifact is individually ready.

## External sources checked

- SLSA source verification expectations: https://slsa.dev/spec/v1.2/verifying-source
- GitHub artifact attestations and subject digest verification: https://docs.github.com/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds
- GitHub `actions/attest` notes that attestations bind a subject plus digest to a predicate: https://github.com/actions/attest

## Why this matters to Ascend

The release claim board is intentionally conservative: the top claims have evidence, but unresolved product, performance, correctness, and release-policy gates still block headline copy. A single fail-closed readiness summary lets future loops check publication eligibility without interpreting Markdown tables or accidentally treating local digests as signed provenance.

## Probe/implementation

- Inspected `fixtures/benchmarks/release-proof-index.ts` and the current release proof index tests.
- Added `ReleaseProofReadinessSummary` to the JSON result.
- Summarized:
  - aggregate `releaseGate`;
  - `headlineClaimsAllowed`;
  - total, missing, and satisfied requirement counts;
  - missing requirements by owner loop;
  - missing requirement IDs by artifact.
- Rendered a `Release Readiness Gate` section in Markdown.
- Kept the aggregate boundary explicit: this is a publication gate over local proof artifacts, not signed provenance or attestation verification.

## Results

Current aggregate readiness:

| Field | Value |
| --- | --- |
| releaseGate | blocked-by-publication-policy |
| headlineClaimsAllowed | false |
| totalRequirementCount | 6 |
| missingRequirementCount | 6 |
| satisfiedRequirementCount | 0 |
| missingByOwnerLoop | correctness=1, performance=1, product=2, release=2 |
| safe-open-proof missing IDs | public-edge-fixtures, release-latency-run, publication-boundary |
| package-action-proof missing IDs | edge-fixture-policy, provenance-boundary, unsupported-feature-boundary |

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/safe-open-proof.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

Focused proof tests passed: 9 pass, 0 fail.

## Confidence

High. The fold-in is additive benchmark/proof metadata, tested at the JSON and Markdown levels, and leaves product APIs, mutation behavior, and proof artifact generation unchanged.

## Fold-in decision

Fold into the release proof harness. Future product/performance and correctness/product owner loops should consume `readiness.headlineClaimsAllowed` and `readiness.missingRequirementCount` as the fail-closed gate before producing release copy.

## Next question

Should the proof owner loops produce separate release-ready Markdown reports only when `readiness.headlineClaimsAllowed` is true, or should they always publish blocked reports with the missing gate list?
