# Compact Report Canonicalization Deferral

Date: 2026-05-15

## Question

Should safe-open and package-action compact report canonicalization share one helper now, or should canonicalization wait until release owners decide the artifact storage policy?

## Hypothesis

Canonicalization should wait. A helper would be easy to add, but it would prematurely encode what counts as the artifact subject before release owners define storage, privacy filtering, and verification expectations.

## External sources checked

- RFC 8785 JSON Canonicalization Scheme: https://www.rfc-editor.org/rfc/rfc8785.html
- SLSA attestation model storage/lookup conventions: https://slsa.dev/spec/v1.1/attestation-model
- SLSA tracks leave provenance distribution and expectations to the ecosystem or organization: https://slsa.dev/spec/latest/levels
- GitHub artifact attestations bind named subjects and digests: https://github.com/actions/attest

## Why this matters to Ascend

Ascend now has compact reports for the top two claims and a release-index pointer to each command. Adding a canonicalization helper before release policy could make a local stable-shape digest look like an approved artifact digest.

## Probe/implementation

I inspected current local canonicalization patterns:

```bash
rg -n "stableJson|canonical|generatedAt|compactReport" fixtures packages apps research/experiments/syntheses/2026-05-release-claim-board.md
```

Findings:

- `fixtures/benchmarks/release-proof-index.ts` has a local `stableJson(stripRunNoise(...))` helper for stable-shape proof digests.
- `fixtures/benchmarks/agent-view-budget-proof.ts`, `agent-view-recovery-proof.ts`, and `journal-law-proof.ts` have benchmark-local stable JSON helpers.
- `packages/io-xlsx/src/writer/index.ts` has a writer-local stable JSON helper for package state comparisons.
- None of these helpers define release artifact subject bytes, privacy filtering, storage location, or verification expectations.

No code was changed.

## Results

The right next move is to keep `compact-report-publication-policy` as the named readiness gate and avoid a shared compact-report canonicalization helper until release owners decide:

- whether compact reports are stored artifacts or generated-on-demand views;
- which fields are excluded from public artifacts;
- whether canonicalization follows RFC 8785, the existing sorted-key helper style, or another format;
- where digests are retained and how users verify them;
- what wording distinguishes local stable-shape hashes from signed provenance.

## Confidence

High. The codebase already has local stable JSON helpers for specific proof and writer contexts, and external references make clear that canonicalization only matters after the artifact subject and verification model are defined.

## Fold-in decision

Archive as a deliberate deferral. Do not implement a shared compact-report canonicalization helper in this block.

## Next question

Should the final claim-steward handoff rank the compact report publication gate ahead of public-edge-fixture replacement, or keep it as a release-loop blocker below product fixture policy?
