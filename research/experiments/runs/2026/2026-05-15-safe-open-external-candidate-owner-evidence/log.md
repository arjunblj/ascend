# Safe Open External Candidate Owner Evidence

## Question

Can the ExcelForge unknown-part workbook candidate become machine-readable owner-review evidence without satisfying the safe-open `public-edge-fixtures` gate?

## Hypothesis

Yes. The candidate should be visible to product/release owners in `release-proof-index --owner-handoffs-json`, but it must carry an explicit gate effect saying it does not satisfy `public-edge-fixtures` because the workbook is not vendored, attribution policy is not approved, and the signed-workbook gap remains open.

## External sources checked

- ExcelForge repository page, which currently shows MIT license evidence and documents XLSX package-signing features: https://github.com/node-projects/excelForge
- ExcelForge raw package manifest, which declares `"license": "MIT"`: https://raw.githubusercontent.com/node-projects/excelForge/master/package.json
- ExcelForge raw workbook candidate: https://raw.githubusercontent.com/node-projects/excelForge/master/src/test/Book%201.xlsx
- Microsoft Protected View documentation, used as the honest competitor boundary for trust UX versus Ascend's package-feature routing: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Open Packaging Conventions overview for package parts and relationships: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

Safe unknown workbook opening is still a top release-claim handoff. The previous external probe improved the unknown-part fixture acquisition story, but the owner handoff did not expose that candidate in the canonical machine-readable release index. Making the candidate visible there helps product/release choose between vendoring a public fixture and accepting generated topology evidence, without letting research silently close the gate.

## Probe/implementation

- Added `externalCandidateEvidence` under `fixturePolicyEvidence.safeOpen` in `fixtures/benchmarks/release-proof-index.ts`.
- Recorded the ExcelForge candidate ID, source URL, MIT package-manifest URL, workbook and manifest SHA-256 values, safe-open routing result, part/relationship counts, sample unknown part, owner decision, gate effect, and boundary.
- Rendered the candidate in the release proof Markdown output.
- Updated generated-fixture decision evidence for `safe-open-proof/unknown-part` to mention the external candidate while saying it awaits owner review and is not vendored.
- Added release-proof-index tests for JSON and Markdown evidence.

## Results

- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json` now reports:
  - `fixturePolicyEvidence.safeOpen.externalCandidateEvidence[0].candidateId=excelforge-book1-unknown-part`;
  - `status=external-candidate-owner-review-required`;
  - `riskFamily=preservedOther`;
  - `recommendedMode=metadata-only`;
  - `reviewBeforeHydration=true`;
  - `partCount=50`;
  - `relationshipCount=37`;
  - `sampleUnknownPart=docMetadata/LabelInfo.xml`;
  - `gateEffect=does-not-satisfy-public-edge-fixtures`.
- The release gate remains fail-closed:
  - `releaseGate=blocked-by-publication-policy`;
  - `headlineClaimsAllowed=false`;
  - `implementationSurfacePromotionAllowed=false`;
  - `missingRequirementCount=9`.
- Validation passed:
  - `bun test fixtures/benchmarks/release-proof-index.test.ts`;
  - `bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`;
  - `bunx tsc --build`.

## Confidence

High that the handoff is more actionable and remains fail-closed. Medium that this candidate will be accepted, because binary vendoring, attribution, and signed-workbook replacement remain owner decisions.

## Fold-in decision

Promote to product/release proof packaging. Do not vendor the workbook and do not mark `public-edge-fixtures` satisfied.

## Next question

Should the package-action `unknown-part-error` gate get an analogous external-candidate owner-review record, or should it stay generated-only because the candidate has not been proven as a fail-closed mutation case?
