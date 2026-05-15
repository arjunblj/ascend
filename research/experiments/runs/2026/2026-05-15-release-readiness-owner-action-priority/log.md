# Release Readiness Owner Action Priority

Date: 2026-05-15

## Question

Should the final claim-steward handoff rank the compact report publication gate ahead of public-edge-fixture replacement, or keep it as a release-loop blocker below product fixture policy?

## Hypothesis

The compact report publication gate should rank below fixture policy and claim-boundary gates. It blocks digest publication and stronger proof packaging, but public/generated fixture policy and correctness boundaries determine whether the product-shaped claims are credible at all.

## External sources checked

- Microsoft Protected View describes read-only review for potentially unsafe files: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- SLSA attestation model separates statement contents from storage/lookup conventions: https://slsa.dev/spec/v1.1/attestation-model
- SLSA tracks leave provenance distribution and expectation policy to the ecosystem or organization: https://slsa.dev/spec/latest/levels
- Open Packaging Conventions describe package signatures over package contents: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

The release proof index now has nine missing gates. Without a prioritized owner queue, loops can spend time on artifact publication mechanics before resolving whether the evidence and boundaries are strong enough to support the claims.

## Probe/implementation

I added `readiness.nextOwnerActions` to `fixtures/benchmarks/release-proof-index.ts`. It ranks missing `readyWhen` gates by proof leverage:

1. fixture disclosure or replacement gates;
2. correctness claim-boundary gates;
3. release latency evidence;
4. streaming wording boundary;
5. publication/provenance wording;
6. compact report publication policy.

The compact report publication gate remains present and fail-closed, but it ranks after evidence and claim-boundary gates.

## Results

The release proof index now emits an ordered owner-action queue. The top actions are:

- `package-action-proof/edge-fixture-policy`
- `safe-open-proof/public-edge-fixtures`
- `package-action-proof/unsupported-feature-boundary`
- `safe-open-proof/release-latency-run`

The compact report publication policy actions rank last because compact report digests are explicitly not promoted yet.

## Confidence

Medium-high. The priority order is a stewardship decision, but it is grounded in current proof blockers and external provenance guidance: storage/canonicalization matters for publication, while fixture and boundary decisions determine the claim itself.

## Fold-in decision

Folded into the benchmark release proof index as machine-readable owner-action priority. No product surface changed.

## Next question

Should the top-ranked product fixture gates be resolved by accepting disclosed generated fixtures, or by adding public binary replacements?
