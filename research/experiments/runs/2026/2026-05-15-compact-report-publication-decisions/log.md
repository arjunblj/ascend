# Compact Report Publication Decisions

## Question

Can the compact report publication blocker be made actionable for release owners without publishing digests or implying attestation?

## Hypothesis

Yes. The release proof index already verifies that compact safe-open and package-action summaries exist and omit forbidden payload fields. The missing release policy can be represented as four pending owner decisions with explicit acceptance evidence and rejection conditions.

## External sources checked

- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations
- SLSA provenance specification: https://slsa.dev/spec/v1.0-rc1/provenance
- in-toto attestations: https://github.com/in-toto/attestation

## Why this matters to Ascend

Ascend’s top two release claims need proof that is shareable and honest. Compact reports are useful owner-review artifacts, but publishing digests without storage, privacy, canonicalization, and verification policy can accidentally imply signed provenance or tamper-evident release evidence. The North Star needs trustworthy claims, not proof-shaped marketing.

## Probe/implementation

- Inspected `compactReportPublicationEvidence` in `fixtures/benchmarks/release-proof-index.ts`.
- Added `policyDecisions` to the compact report publication evidence:
  - artifact storage path
  - retention and privacy filtering
  - canonicalization subject
  - offline verification expectations
- Rendered those decisions in release-proof Markdown and owner-handoff JSON.
- Added regression assertions in `fixtures/benchmarks/release-proof-index.test.ts`.

## Results

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
```

Result: 4 tests passed.

Current owner evidence:

| Field | Value |
| --- | --- |
| Compact report digests indexed | `false` |
| Forbidden payload fields embedded | `false` |
| GeneratedAt included | `true` |
| Policy decisions | `artifact storage path`, `retention and privacy filtering`, `canonicalization subject`, `offline verification expectations` |
| Owner gate | `compact-report-publication-policy` remains missing for both top artifacts |

## Confidence

High for the handoff shape and validation coverage. Medium for release readiness: an owner still needs to define and approve the actual publication policy before compact reports can be treated as release artifacts.

## Fold-in decision

Promote to release loop as owner-decision evidence only. Do not publish compact report digests, do not call compact reports signed provenance, and do not satisfy `compact-report-publication-policy`.

## Next question

Can the release proof index make correctness-boundary failures stable and diagnostic if any unsupported-feature evidence regresses, instead of relying on broad `allCurrentEvidencePresent` expectations?
