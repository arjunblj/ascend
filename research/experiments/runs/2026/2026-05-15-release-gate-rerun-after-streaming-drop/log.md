# Release Gate Rerun After Streaming Drop Proof

## Question

After the streaming `drop` proof and in-flight journal validation commits, do the top release claims become publishable, or do owner gates still block headline wording and new implementation surfaces?

## Hypothesis

The proof improves package-action evidence but should not change release posture. The release index should still block headline claims because fixture policy, latency policy, provenance wording, compact report publication, and owner approvals remain unresolved.

## External sources checked

- SLSA provenance v1.0-rc1: https://slsa.dev/spec/v1.0-rc1/provenance
- GitHub artifact attestations provenance docs: https://docs.github.com/en/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds

## Why this matters to Ascend

The North Star needs proof-backed claims, not accidental marketing drift. A stronger local package-action matrix is useful only if the release gate keeps unsupported wording out: no signed provenance, no malware safety, no full streaming parity, and no new public surfaces from research alone.

## Probe/implementation

Ran the current owner-handoff proof index from a tracked-clean state except long-lived untracked research/tmp files:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
```

## Results

- `releaseGate=blocked-by-publication-policy`
- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- Top implementation handoffs remain exactly two:
  - `safe-open-proof`
  - `package-action-proof`
- Package-action performance gate now records representative streaming proof for `passthrough/regenerate/add/drop`.
- Streaming `error`, public macro, and public chart cases remain outside streaming parity.
- SLSA/GitHub attestation references remain boundary checks only; Ascend still produces local proof summaries, not signed provenance.

## Confidence

High. The result is machine-readable release-index output and matches the intended no-promotion posture.

## Fold-in decision

Promote to topic synthesis and owner handoff only. Do not fold into production. Do not add SDK, CLI, API, or MCP surfaces. Keep formula rename frozen and keep lower-ranked portfolio items out of implementation handoff.

## Next question

Should the next implementation owner tackle product acceptance of disclosed generated fixtures, or release-owned compact report publication/canonicalization policy?
