# Release Index Compact Report Commands

Date: 2026-05-15

## Question

Can the release proof index expose compact report commands for the top two claims without indexing compact report digests or implying artifact attestation?

## Hypothesis

Yes. The release index can name compact report commands as reproduction pointers while leaving digest storage out of scope. This helps owner loops find the claim-safe handoff artifacts without creating a false provenance or storage claim.

## External sources checked

- GitHub artifact attestations bind named subjects and digests to predicates: https://github.com/actions/attest
- GitHub artifact attestation docs distinguish subject paths and subject digests: https://docs.github.com/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds
- SLSA verification expects provenance subjects to match artifact digests: https://slsa.dev/spec/v1.2/verifying-artifacts
- Microsoft OPC digital signatures validate signed package contents: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

The top release claims now have compact report generators, but prior digest-policy probes showed raw compact JSON digests churn because `generatedAt` is included. The release index should make compact reports discoverable while keeping artifact storage, privacy, and canonicalization policy explicit and unpromoted.

## Probe/implementation

I added an optional `compactReportCommand` field to release proof index artifacts and populated it for:

- `safe-open-proof`: `bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json`
- `package-action-proof`: `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json`

The index still records digests only for the existing full proof artifacts. It does not store compact report bytes, compact report digests, signed provenance, or attestation status.

## Results

The release index tests now assert both compact report commands are present in JSON and Markdown output. The existing fail-closed release gate remains unchanged:

- `headlineClaimsAllowed=false`
- `releaseGate=blocked-by-publication-policy`
- compact report digests remain excluded

## Confidence

High. This is a small discoverability change with tests, and it preserves the digest policy boundary from the prior two compact-report probes.

## Fold-in decision

Folded into the benchmark release proof index. Do not promote compact report digests until artifact storage, privacy filtering, and stable canonicalization are owner-approved.

## Next question

Should the release-claim board make compact report publication policy a named proof requirement for both top claims, or is the current do-not-promote boundary enough?
