# Fixture Acquisition Owner Ranking

## Question

Can Ascend rank the remaining public fixture acquisition tasks for the top release claims so product/release owners know what to prove next?

## Hypothesis

Yes. Current release proof evidence already identifies which generated structural cases remain and which external candidates exist. A small owner-handoff plan can rank the next fixture decisions without changing workbook behavior or treating owner-review candidates as gate satisfaction.

## External sources checked

- Microsoft Open Packaging Conventions overview and digital signature concepts: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- OOXML Digital Signature Origin Part reference: https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Digital_topic_ID0EGZAO.html
- Microsoft Open XML SDK `DigitalSignatureOriginPart`: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.digitalsignatureoriginpart
- ExcelForge repository and README: https://github.com/node-projects/excelForge
- ExcelForge sample workbook: https://raw.githubusercontent.com/node-projects/excelForge/master/src/test/Book%201.xlsx

## Why this matters to Ascend

The North Star requires credible release claims, not an ever-growing list of experiments. Fixture gaps are currently a top blocker for safe unknown workbook opening and auditable package-part mutation. Ranking the acquisition tasks turns the claim board into owner action: review the shared unknown-part candidate first, then decide signed-package evidence, then malformed-package policy.

## Probe/implementation

Implemented `fixtureAcquisitionPlan` in `fixtures/benchmarks/release-proof-index.ts`:

- rank 1: review/vendoring decision for the shared ExcelForge unknown-part candidate.
- rank 2: signed XLSX package fixture acquisition or explicit generated-topology acceptance.
- rank 3: malformed package generated-bytes policy decision.

The plan is emitted in:

- full release proof index result.
- compact owner-handoff JSON.
- human-readable release proof Markdown.

The plan carries task, already-present evidence, proof still missing, validation command, reference, kill criterion, owner decision, and honest boundary for each row.

## Results

Validation:

- `bunx biome check --write fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`
- `bun test fixtures/benchmarks/release-proof-index.test.ts`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`
- `bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`
- `bunx tsc --build`

The owner-handoff JSON now includes:

- `fixtureAcquisitionPlan.status=ranked-owner-review-required`
- `taskCount=3`
- top task `unknown-part-shared-candidate` related to both `safe-open-proof` and `package-action-proof`
- signed-package task preserving the signature validation/trust boundary
- malformed-package task preserving fail-closed rejection-only wording

Release gates remain blocked:

- `releaseGate=blocked-by-publication-policy`
- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`

## Confidence

High for owner routing. Medium for the exact ranking, because owners may value signed-package acquisition above unknown-part candidate review. The ranking is still defensible because the unknown-part candidate is already identified and could reduce uncertainty for both top claims.

## Fold-in decision

Fold into release proof owner-handoff evidence only. This is a claim stewardship artifact, not a product surface. Do not promote stronger safe-open or package-action wording until owner decisions are made and tracked fixtures or generated-fixture acceptance policy exists.

## Next question

Can the next owner-facing proof distinguish which release blockers are decision-only versus validation-run blockers, and produce separate prompts for product, correctness, performance, and release loops?
