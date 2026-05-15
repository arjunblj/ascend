# Compact Report Publication ReadyWhen

Date: 2026-05-15

## Question

Should the release-claim board make compact report publication policy a named proof requirement for both top claims, or is the current do-not-promote boundary enough?

## Hypothesis

The policy should be a named `readyWhen` requirement. Once the release proof index exposes compact report commands, artifact storage and canonicalization are no longer just editorial concerns; they are release blockers for stronger proof wording.

## External sources checked

- SLSA tracks leave provenance distribution and expectation policy to the ecosystem or organization: https://slsa.dev/spec/latest/levels
- SLSA provenance identifies artifacts with digests: https://slsa.dev/spec/v1.0-rc1/provenance
- SLSA attestation model names storage/lookup conventions as part of attestation use: https://slsa.dev/spec/v1.1/attestation-model
- GitHub artifact attestations bind named subjects and digests to predicates: https://github.com/actions/attest

## Why this matters to Ascend

Ascend's top two claims now have compact reports and release-index command pointers. Without a named readiness gate, future release work could accidentally publish compact report digests without deciding artifact storage, privacy filtering, or stable canonicalization.

## Probe/implementation

I added `compact-report-publication-policy` as a missing `readyWhen` requirement to both top artifacts in `fixtures/benchmarks/release-proof-index.ts`.

Owner loop: `release`

Requirement:

```text
define artifact storage, privacy filtering, and canonicalization policy before compact report digests are indexed or published
```

## Results

The release readiness summary now fails closed with:

- total requirements: 9
- missing requirements: 9
- release-owned missing requirements: 4
- `safe-open-proof` missing gates include `compact-report-publication-policy`
- `package-action-proof` missing gates include `compact-report-publication-policy`

The compact report commands remain discoverability pointers, not digested release artifacts.

## Confidence

High. The prior two digest probes showed why raw compact report digests should not be promoted, and external provenance systems require explicit subject/storage/verification semantics.

## Fold-in decision

Folded into the benchmark release proof index as a named release-loop gate.

## Next question

Should safe-open and package-action compact report canonicalization share one helper, or should canonicalization wait until release owners decide the artifact storage policy?
