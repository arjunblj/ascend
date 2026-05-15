# Safe Open Owner Acceptance Checklist

## Question

Can safe-open release wording be made owner-approvable by turning the public/generated fixture and latency blockers into exact acceptance checkboxes, without rerunning the entire proof matrix?

## Hypothesis

Yes. The current safe-open proof and release index already identify the blockers. The next useful artifact is an owner checklist, not another open surface.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Office File Validation and Protected View: https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/office-file-fails-validation
- Microsoft Trusted Documents and active content: https://support.microsoft.com/en-us/office/trusted-documents-cf872bd8-47ec-4c02-baa5-1fdba1a11b53
- Application Guard for Office admin docs: https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/install-app-guard

## Why this matters to Ascend

Safe unknown workbook opening is the top product/performance handoff. The claim is credible only if generated fixtures and latency wording are explicitly accepted; otherwise the product could overclaim trust, sandboxing, active-content safety, or performance thresholds.

## Probe/implementation

- Ran a local Bun probe over `runSafeOpenProof`, `safeOpenCompactReleaseReport`, and `runReleaseProofIndex`.
- Confirmed current safe-open proof shape and missing `readyWhen` gates.
- Updated `research/experiments/syntheses/2026-05-owner-handoff.md` with safe-open acceptance checkboxes.
- Updated `research/experiments/syntheses/2026-05-release-claim-board.md` with the same owner-ready boundary.

## Results

Current proof input:

| Field | Value |
| --- | --- |
| Proof cases | 9 |
| Public fixture cases | 6 |
| Generated edge-package cases | 2 |
| Malformed cases | 1 |
| Generated/disclosed cases | `signed`, `unknown-part`, `malformed` |
| Review before hydration | 4 |
| Malformed rejected | `true` |
| Headline claim allowed | `false` |
| Release gate | `blocked-by-publication-policy` |

Acceptance checklist:

- Product accepts disclosed generated `signed` and `unknown-part` structural fixtures, or replaces them with public binary fixtures.
- Product accepts generated malformed-package evidence as structural rejection proof, not real-world repair behavior.
- Performance runs tracked-clean release-environment latency over standardized public inputs and approves wording that does not read as a threshold unless thresholds are explicitly owned.
- Release approves boundary wording: not malware scanning, sandboxing, file trust, active-content safety, signed provenance, or malformed-package recovery.
- Release keeps compact report digests unpublished until storage, privacy filtering, canonicalization, and verification expectations exist.

## Confidence

High that the checklist reflects current proof. Medium that the claim can be released, because owner approval remains outside research.

## Fold-in decision

Promote to topic synthesis only. Keep `safe-open-proof` blocked until product, performance, and release owners complete the checklist. Do not add new open surfaces.

## Next question

Can package-action owner acceptance be reduced to exact checkboxes in the same style, covering edge fixtures, unsupported features, streaming wording, provenance, and compact report publication?
