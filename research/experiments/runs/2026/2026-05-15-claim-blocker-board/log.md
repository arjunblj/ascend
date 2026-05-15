# Claim Blocker Board

## Question

Can the release proof index expose a compact claim blocker board grouped by product-shaped claim and owner loop?

## Hypothesis

Yes. The index already has canonical missing `readyWhen` gates and ranked owner actions. A derived board can group those actions by claim and owner without becoming a second source of truth.

## External sources checked

- GitHub required status checks: https://docs.github.com/repositories/configuring-branches-and-merges-in-your-repository/managing-protected-branches/about-protected-branches
- GitHub artifact attestation verification: https://docs.github.com/actions/concepts/security/artifact-attestations
- SLSA provenance specification: https://slsa.dev/spec/v1.0-rc1/provenance

## Why this matters to Ascend

Ascend’s top claims are blocked by different owner loops. Product, correctness, performance, and release owners should not have to reconstruct their work from multiple proof objects. A claim/owner blocker board keeps the release workflow focused on proving product-shaped claims rather than promoting random surfaces.

## Probe/implementation

- Inspected `releaseReadinessSummary`, `nextOwnerActions`, and `implementationHandoffs`.
- Added `readiness.claimBlockerBoard` and owner-handoff JSON output.
- The board is derived from missing `readyWhen` gates and ranked owner actions.
- Rendered the board in release-proof Markdown.
- Added assertions for safe-open and package-action owner rows in `fixtures/benchmarks/release-proof-index.test.ts`.

## Results

Current board rows:

| Claim | Owner rows |
| --- | --- |
| `safe-open-proof` | performance, product, release |
| `package-action-proof` | correctness, performance, product, release |

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
```

Result: 4 tests passed.

## Confidence

High for consistency because the board is derived from canonical missing gates and owner actions. Medium for ergonomics because owner loops may still want a shorter human report, but the data shape now exists.

## Fold-in decision

Promote to product, correctness, performance, and release loops as owner routing evidence. This is not a new product surface, not a satisfied gate, and not permission to publish any claim.

## Next question

Can safe-open release latency evidence be packaged as an owner-approved validation run over public inputs without turning local timing into a release threshold?
