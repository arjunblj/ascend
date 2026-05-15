# Compact Report Publication Policy Audit

## Question

Can the compact-report publication gate be reduced to an owner policy decision, or is there a real missing proof around artifact storage, privacy filtering, and canonicalization?

## Hypothesis

It should remain a release-owner policy gate. The compact reports are already minimized local summaries and command pointers, but publishing them by digest would require storage, privacy filtering, and canonicalization decisions that research should not invent.

## External sources checked

- RFC 8785 JSON Canonicalization Scheme: https://www.ietf.org/rfc/rfc8785.html
- GitHub offline attestation verification docs: https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/verify-attestations-offline
- SLSA specification v1.2: https://slsa.dev/spec/v1.2/
- W3C Privacy Principles: https://www.w3.org/TR/privacy-principles/

## Why this matters to Ascend

The top release claims need reproducible proof, but publishing compact reports without policy could imply stable artifact semantics, privacy review, or verification guarantees that Ascend has not defined.

## Probe/implementation

- Inspected `fixtures/benchmarks/release-proof-index.ts`, `fixtures/benchmarks/safe-open-proof.ts`, and `fixtures/benchmarks/package-action-proof.ts`.
- Ran a local Bun audit over `runReleaseProofIndex`, `safeOpenCompactReleaseReport`, and `packageActionCompactReleaseReport`.
- Confirmed compact report commands exist for both top artifacts.
- Confirmed both top artifacts keep `compact-report-publication-policy` missing and release-owned.
- Confirmed the release index does not include compact digest fields.
- Confirmed compact reports do not embed workbook bytes or full proof artifacts.
- Updated `research/experiments/syntheses/2026-05-release-claim-board.md` with the publication policy boundary.

## Results

Local audit:

| Check | Safe open | Package action |
| --- | --- | --- |
| Compact report command present | yes | yes |
| `compact-report-publication-policy` readyWhen present | yes | yes |
| Compact digest indexed | no | no |
| Compact report embeds workbook bytes | no | no |
| Compact JSON bytes | 3755 | 4258 |

Compact report top-level shapes:

| Artifact | Top-level fields |
| --- | --- |
| Safe open | `boundary`, `caseKindCounts`, `cases`, `claim`, `command`, `coverage`, `generatedAt`, `headlineClaimAllowed`, `publicationStatus`, `readyWhen`, `releaseGate` |
| Package action | `boundary`, `cases`, `claim`, `combinedCommitActionCounts`, `command`, `coverage`, `generatedAt`, `headlineClaimAllowed`, `publicationStatus`, `readyWhen`, `releaseGate`, `sourceCaseCounts` |

Boundary matrix:

| Boundary | Allowed claim | Forbidden claim |
| --- | --- | --- |
| Compact command pointer | compact report commands reproduce claim-safe summaries locally | compact reports are published release evidence artifacts |
| Compact digest indexing | digest publication is deferred until storage and canonicalization policy exists | compact report digests are stable release commitments |
| Workbook byte minimization | compact reports are minimized summaries | compact reports are privacy-reviewed artifacts |
| Canonicalization | full proof artifacts have stable-shape digests | compact report bytes are canonical or signer-ready |

## Confidence

High that current reports are minimized and fail closed. Medium that they are safe to publish, because release policy still needs to define artifact storage, retention/privacy filtering, canonicalization, and verification expectations.

## Fold-in decision

Promote to topic synthesis only. Keep `compact-report-publication-policy` missing for both top artifacts. Do not add compact digest publication or canonicalization helpers until a release owner defines the artifact subject and storage policy.

## Next question

Can the ranked portfolio be reduced into a one-page owner handoff that names only the top two implementation loops and explicitly freezes the remaining 6-8 directions as do-not-promote?
