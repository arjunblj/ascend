# Owner Handoff JSON Source Refresh

## Question

Can the human owner handoff point directly at machine-readable release proof fields, instead of duplicating the claim ladder and deferred-claim list in prose?

## Hypothesis

Yes. The release proof index now carries `implementationHandoffs`, `proofRequired`, `implementationSurfacePromotionAllowed`, `deferredClaims`, and `excludedEvidence`. Refreshing the owner handoff around those fields should reduce drift and keep owner loops tied to the canonical JSON artifact.

## External sources checked

- in-toto defines attestations as authenticated metadata about software artifact subjects and predicates; Ascend's local proof index should be treated as local metadata, not an attestation: https://github.com/in-toto/attestation/blob/main/spec/README.md
- SLSA provenance documents downstream verification expectations around build metadata; this reinforces that proof fields need explicit verification and owner approval: https://slsa.dev/spec/v1.0-rc1/provenance
- GitHub artifact attestations describe cryptographically signed provenance claims and verification; Ascend's release proof index intentionally remains below that boundary: https://docs.github.com/en/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

Owner loops need clear next actions, but Markdown checklists drift if they become a second source of truth. The North Star is credible release claims; the owner handoff should tell product/correctness/performance/release owners which JSON fields to consume, which claims are top handoffs, and which directions are deferred.

## Probe/implementation

Updated `research/experiments/syntheses/2026-05-owner-handoff.md` to:

- mention `implementationSurfacePromotionAllowed=false`;
- note that practical latency public-tracked evidence now uses generated edit input and remains excluded from release proof;
- point owners at `readiness.implementationHandoffs`, `proofRequired`, `deferredClaims`, and `excludedEvidence`;
- align the "Do Not Promote" table with the machine-readable deferred claim statuses.

Probe command:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

## Results

- Release proof index currently reports:
  - `releaseGate=blocked-by-publication-policy`;
  - `implementationSurfacePromotionAllowed=false`;
  - 2 implementation handoffs;
  - 6 deferred claims;
  - 1 excluded evidence item.
- The owner handoff now describes how to consume those fields instead of adding another machine-checked Markdown acceptance table.

## Confidence

High for synthesis alignment. Medium for owner-loop adoption until the product/performance/correctness/release prompts are run directly from the JSON fields.

## Fold-in decision

Fold into synthesis only. No production surface and no proof harness change.

## Next question

Can the safe-open owner resolve the first product gate by explicitly accepting generated structural fixtures, or does the release proof need public binary signed/unknown workbooks?
