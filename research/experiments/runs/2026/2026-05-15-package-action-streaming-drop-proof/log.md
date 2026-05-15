# Package Action Streaming Drop Proof

## Question

Can the existing package-action proof harness cover `drop` in streaming mode without adding a new SDK, CLI, API, or MCP surface, and should it still refuse streaming `error` wording?

## Hypothesis

Yes for calc-chain `drop` if the streaming proof path carries the operation-level recalculation signal into the writer. No for unknown-part `error`; that should remain a non-streaming proof boundary until the writer has deliberate streaming error evidence.

## External sources checked

- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- ECMA-376 Open XML standards page, including Part 2 Open Packaging Conventions: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- GitHub artifact attestations provenance docs: https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
- in-toto attestation spec: https://github.com/in-toto/attestation/blob/main/spec/README.md

## Why this matters to Ascend

The product-shaped claim is auditable package-part mutation: Ascend should explain workbook writes as `passthrough`, `regenerate`, `add`, `drop`, or `error` without implying signed provenance, full streaming parity, or semantic support for every package feature. Covering `drop` in the streaming proof matrix makes the performance boundary more honest and more useful.

## Probe/implementation

- Inspected `fixtures/benchmarks/package-action-proof.ts`, `fixtures/benchmarks/release-proof-index.ts`, and related tests.
- Tried enabling streaming probes for both `calc-chain-drop` and `unknown-part-error`.
- The first probe failed for `calc-chain-drop` because the streaming proof path passed dirty sheet names but not `calcChainDirty`.
- Folded in a benchmark-harness fix: `runStreamingPackageActionProof` now passes `calcChainDirty: applied.value.recalcRequired`.
- Kept `unknown-part-error` non-streaming after the probe still failed to produce the expected streaming `error` action.
- Updated release proof index evidence, owner handoff wording, release claim board, and ranked portfolio to say representative streaming proof covers `passthrough/regenerate/add/drop`, not `error`, macro, chart, or full parity.

## Results

- `package-action-proof` compact coverage now reports `streamingProofCases=3` and `streamingRegenerateParts=2`.
- Streaming proof cases are `docprops-passthrough`, `add-sheet-part`, and `calc-chain-drop`.
- Streaming covered action kinds are `passthrough`, `regenerate`, `add`, and `drop`.
- Streaming missing action kind is `error`.
- `release-proof-index --owner-handoffs-json` remains fail-closed: `releaseGate=blocked-by-publication-policy`, `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, and `missingRequirementCount=9`.
- `streaming-matrix-boundary` remains owner-approval required.

## Confidence

High for the narrow proof-harness improvement and boundary wording. Medium for broader streaming parity, because the unknown-part error case, macro fixture, and chart fixture are still non-streaming in the release proof matrix.

## Fold-in decision

Promote to performance loop as a benchmark/proof-harness fold-in only. Do not promote a new production surface and do not claim full streaming parity. Keep the release wording to: representative streaming proofs cover `passthrough/regenerate/add/drop`; streaming `error`, macro, and chart coverage remain unproven.

## Next question

Should performance owners accept this representative streaming matrix for narrow release wording, or explicitly fund a streaming `error` proof and public macro/chart streaming probes before any release mention?
