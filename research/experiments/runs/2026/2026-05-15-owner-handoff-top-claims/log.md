# Owner Handoff Top Claims

## Question

Can the ranked portfolio be reduced into a one-page owner handoff that names only the top two implementation loops and freezes the remaining directions as do-not-promote?

## Hypothesis

Yes. The current release proof index already limits claim proof to two artifacts, and the owner action ranking is enough to define the next implementation loops without adding new surfaces.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Open Packaging Conventions fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl tutorial preservation boundary: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

Research has produced many plausible directions. The current highest-value synthesis is deciding what Ascend should prove next, not adding another narrow surface because it is available.

## Probe/implementation

- Ran a local Bun probe over `runReleaseProofIndex({ includeTimings: false })`.
- Confirmed only two release artifacts are indexed: `safe-open-proof` and `package-action-proof`.
- Confirmed `headlineClaimsAllowed=false` and 9 missing requirements remain.
- Extracted the first six owner actions from the release readiness queue.
- Added `research/experiments/syntheses/2026-05-owner-handoff.md`.

## Results

Handoff probe:

| Field | Value |
| --- | --- |
| Release artifact count | 2 |
| Top artifacts | `safe-open-proof`, `package-action-proof` |
| Headline claims allowed | `false` |
| Missing requirements | 9 |

Top owner actions:

| Rank | Artifact | Requirement | Owner | Priority |
| ---: | --- | --- | --- | --- |
| 10 | `package-action-proof` | `edge-fixture-policy` | product | claim-evidence |
| 10 | `safe-open-proof` | `public-edge-fixtures` | product | claim-evidence |
| 20 | `package-action-proof` | `unsupported-feature-boundary` | correctness | claim-boundary |
| 30 | `safe-open-proof` | `release-latency-run` | performance | claim-evidence |
| 40 | `package-action-proof` | `streaming-matrix-boundary` | performance | claim-boundary |
| 50 | `package-action-proof` | `provenance-boundary` | release | publication-policy |

## Confidence

High that the top two handoffs are correct for this block. Medium that owner loops will accept generated fixture and publication policy boundaries without additional evidence.

## Fold-in decision

Promote to topic synthesis only. Hand off safe unknown workbook opening and auditable package-part mutation. Keep formula rename, columnar sidecars, release proof bundle attestation, retained viewport collaboration claims, and exact-token agent-view wording frozen.

## Next question

Can safe-open release wording be made owner-approvable by turning the public/generated fixture and latency blockers into exact acceptance checkboxes, without rerunning the entire proof matrix?
