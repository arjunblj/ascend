# Compact Report Publication Evidence

## Question

Can the release owner handoff prove the current compact reports are claim-safe local summaries while still blocking compact report digest publication?

## Hypothesis

Yes. The compact reports already omit workbook bytes and proof digest payload fields. The release proof index can expose that evidence directly, while still requiring owner policy for storage, privacy filtering, canonicalization, and verification.

## External sources checked

- RFC 8785 JSON Canonicalization Scheme: https://www.rfc-editor.org/rfc/rfc8785
- in-toto attestation specification: https://github.com/in-toto/attestation/tree/main/spec
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations
- SLSA provenance: https://slsa.dev/spec/v1.0-rc1/provenance

## Why this matters to Ascend

Ascend should be able to give release owners compact proof summaries without implying stable published artifacts, signed provenance, tamper evidence, or canonical JSON bytes. This keeps release proof useful while preventing accidental overclaiming.

## Probe/implementation

Added `compactReportPublicationEvidence` to `fixtures/benchmarks/release-proof-index.ts` and the compact owner handoff JSON. It derives evidence from the current safe-open and package-action compact reports:

- compact report command presence
- compact JSON byte size
- top-level field inventory
- forbidden payload field scan
- release-index `compact-report-publication-policy` gate presence
- `generatedAt` presence
- missing publication policy requirements

Forbidden payload fields include workbook/proof byte and digest-like fields such as `inputBytes`, `outputBytes`, `inputSha256`, `sourceSha256`, `outputSha256`, `sha256`, `stableShapeSha256`, `proofJsonBytes`, and `streamingRegeneratePartPaths`.

## Results

Probe command:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json | jq '{headlineClaimsAllowed, implementationSurfacePromotionAllowed, missingRequirementCount, compactReportPublicationEvidence}'
```

Observed:

- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- `compactReportPublicationEvidence.status=local-summary-present-publication-policy-required`
- `compactReportPublicationEvidence.compactReportDigestsIndexed=false`
- `compactReportPublicationEvidence.allCompactCommandsPresent=true`
- `compactReportPublicationEvidence.compactReportsEmbedForbiddenPayloadFields=false`
- `compactReportPublicationEvidence.generatedAtIncluded=true`
- both reports have `readyWhenGatePresent=true`
- missing requirements remain: artifact storage path, retention/privacy filtering, canonicalization subject, offline verification expectations

Targeted validation:

- `bun test fixtures/benchmarks/release-proof-index.test.ts`

## Confidence

High that the compact report privacy/canonicalization boundary is now machine-visible. Medium for publication readiness, because release owners still need to define storage and verification semantics.

## Fold-in decision

Folded into release proof indexing as release-owner evidence. Do not index compact report digests, publish compact reports as release artifacts, or claim canonical/signed provenance until release policy exists.

## Next question

Should the next loop generate a release-owner prompt for compact report storage/canonicalization policy, or move to performance-owned release latency evidence for safe-open?
