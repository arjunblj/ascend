# Compact Report Field Privacy Inventory

## Question

Can compact report privacy review be reduced to field-level inventory rather than adding canonicalization or storage policy?

## Hypothesis

Yes for diagnostics, no for publication. A field-level inventory can prove the compact reports omit workbook bytes and full proof artifacts, but it cannot replace release-owned storage, privacy filtering, canonicalization, and verification policy.

## External sources checked

- W3C privacy protection principles: https://www.w3.org/wiki/Privacy/Privacy_protection_principles
- NIST Privacy Framework: https://www.nist.gov/privacy-framework
- RFC 8785 JSON Canonicalization Scheme: https://www.rfc-editor.org/rfc/rfc8785.html
- GitHub artifact attestation offline verification: https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/verify-attestations-offline

## Why this matters to Ascend

Compact reports are useful for claim review, but publishing them prematurely could imply privacy review, stable byte canonicalization, or offline verification that Ascend has not defined.

## Probe/implementation

Ran a local Bun inventory over:

- `safeOpenCompactReleaseReport(await runSafeOpenProof({ includeTimings: false }))`
- `packageActionCompactReleaseReport(await runPackageActionProof({ includeTimings: false }))`

The probe listed top-level fields, checked for banned artifact fields, and captured path-like strings.

## Results

| Artifact | JSON bytes | Top-level fields | Banned artifact fields | Path-like strings |
| --- | ---: | --- | --- | --- |
| Safe open compact report | 3755 | `boundary`, `caseKindCounts`, `cases`, `claim`, `command`, `coverage`, `generatedAt`, `headlineClaimAllowed`, `publicationStatus`, `readyWhen`, `releaseGate` | none | command path plus public fixture paths |
| Package action compact report | 4258 | `boundary`, `cases`, `claim`, `combinedCommitActionCounts`, `command`, `coverage`, `generatedAt`, `headlineClaimAllowed`, `publicationStatus`, `readyWhen`, `releaseGate`, `sourceCaseCounts` | none | command path plus public fixture paths |

Banned terms checked:

- `inputSha256`
- `outputBytes`
- `proofJsonBytes`
- `sourceByteDigests`
- `outputByteDigests`
- `workbookBytes`
- `artifactBytes`

## Confidence

High that the compact reports omit workbook bytes and full proof artifacts today. Medium that they are privacy-safe for publication, because path disclosure and future field additions still need release policy.

## Fold-in decision

Promote to topic synthesis only. Do not add compact report digests, canonicalization helpers, or publication surfaces. Keep `compact-report-publication-policy` missing until release owners define storage, privacy filtering, canonicalization, and verification expectations.

## Next question

Can this block end with a compact handoff summary, or is there another owner-blocking diagnostic that must be run before handing back?
