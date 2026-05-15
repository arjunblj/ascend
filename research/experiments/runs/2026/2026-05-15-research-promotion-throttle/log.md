# Research Promotion Throttle

## Question

Can the next research cycle safely move away from release-claim stewardship and return to a single performance unknown, or should owner-loop approval remain the blocker before new experiments?

## Hypothesis

Owner-loop approval should remain the promotion blocker. Research can still run diagnostic probes, but it should not promote new product claims or public surfaces while the top two release claims have unresolved owner gates.

## External sources checked

- Google SRE launch checklist: https://sre.google/sre-book/launch-checklist/
- Google SRE reliable product launches: https://sre.google/sre-book/reliable-product-launches/
- SLSA verified properties and attestation model: https://slsa.dev/spec/v1.1/attestation-model
- Stage-gate overview: https://www.stage-gate.com/blog/the-stage-gate-model-an-overview/

## Why this matters to Ascend

The research loop has already identified the next two highest-leverage claims. Adding new surfaces before owner approval would create claim debt and compete with correctness, performance, and product loops.

## Probe/implementation

- Ran a local Bun probe over `runReleaseProofIndex({ includeTimings: false })`.
- Read current `nextOwnerActions` and headline gate state.
- Updated `research/experiments/syntheses/2026-05-owner-handoff.md` with a promotion throttle rule.

## Results

| Field | Value |
| --- | --- |
| Headline claims allowed | `false` |
| Missing owner requirements | 9 |
| Decision | hold new promotion; owner approval remains blocker |

Blocking owner actions:

- rank 10: `package-action-proof/edge-fixture-policy` owned by product
- rank 10: `safe-open-proof/public-edge-fixtures` owned by product
- rank 20: `package-action-proof/unsupported-feature-boundary` owned by correctness
- rank 30: `safe-open-proof/release-latency-run` owned by performance
- rank 40: `package-action-proof/streaming-matrix-boundary` owned by performance
- rank 50: `package-action-proof/provenance-boundary` owned by release
- rank 50: `safe-open-proof/publication-boundary` owned by release
- rank 60: `package-action-proof/compact-report-publication-policy` owned by release
- rank 60: `safe-open-proof/compact-report-publication-policy` owned by release

## Confidence

High. The release gate is already machine-readable and says the portfolio should hold promotion.

## Fold-in decision

Promote to topic synthesis only. Keep research in diagnostic/stewardship mode until owner loops clear or explicitly reject the top claim gates.

## Next question

Which diagnostic probe is most useful while promotion is throttled: latency input provenance validation, public fixture search, or compact report privacy review?
