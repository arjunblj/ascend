# Package Action External Candidate Owner Evidence

## Question

Can the release proof index expose a real external workbook candidate for the package-action unknown-part error claim without pretending it satisfies the tracked edge-fixture policy?

## Hypothesis

Yes. The ExcelForge `Book 1.xlsx` sample contains `docMetadata/LabelInfo.xml`, which Ascend classifies as an unknown preserved package part. A local mutation probe should show package-action proof can explain the workbook as passthrough/regenerate/error evidence, while the release gate remains blocked because the workbook is not vendored or owner-approved.

## External sources checked

- ExcelForge repository and README: https://github.com/node-projects/excelForge
- ExcelForge sample workbook: https://raw.githubusercontent.com/node-projects/excelForge/master/src/test/Book%201.xlsx
- ExcelForge package manifest license evidence: https://raw.githubusercontent.com/node-projects/excelForge/master/package.json
- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- SheetJS issue tracker for real-world XLSX preservation/corruption pressure: https://git.sheetjs.com/sheetjs/sheetjs/issues
- openpyxl documentation index for supported workbook IO scope: https://openpyxl.readthedocs.io/

## Why this matters to Ascend

Ascend's auditable package-part mutation claim is strongest when the proof board can point to real workbook shapes, not only generated structural packages. The North Star is not "mutate everything"; it is preservation-first, explainable workbook mutation. This candidate helps owners decide whether a real public unknown-part workbook can replace or supplement the generated unknown-part error fixture.

## Probe/implementation

Local probe:

```bash
bun --eval 'import { createAgentPlan, commitAgentPlan, createAgentCommitPackageActionProof } from "./packages/sdk/src/index.ts"; const input="/private/tmp/excelForge-Book1.xlsx"; const output="/private/tmp/excelForge-package-action-probe-output.xlsx"; const ops=[{op:"setCells", sheet:"Projekt 1", updates:[{ref:"A1", value:"probe"}]}]; const plan=await createAgentPlan(input, ops); const commit=await commitAgentPlan(input, ops, {output, approvals:"all", allowLoss:"all"}); const proof=createAgentCommitPackageActionProof(commit); const unknown=proof.actions.filter(a => (a.partPath ?? "").includes("docMetadata") || (a.partPath ?? "").includes("LabelInfo")); console.log(JSON.stringify({planOk: plan.writePolicy.ok, commitOk: commit.writePolicy.ok, auditsPassed: commit.postWrite.auditsPassed, byAction: proof.byAction, issueCount: proof.issues.length, unknown, packageIssueRefs: (commit.apply.journal?.issues ?? []).filter(i=>i.surface==="package-parts").flatMap(i=>i.refs ?? [])}, null, 2));'
```

Implementation:

- Added `fixturePolicyEvidence.packageAction.externalCandidateEvidence` to `fixtures/benchmarks/release-proof-index.ts`.
- Added a Markdown row for package-action external candidates in the release proof index.
- Added owner-handoff and Markdown regression assertions in `fixtures/benchmarks/release-proof-index.test.ts`.
- Kept `gateEffect=does-not-satisfy-edge-fixture-policy`.

## Results

The local probe produced:

- `planOk=false`
- `commitOk=false`
- `postWriteAuditsPassed=false`
- action counts: `passthrough=42`, `regenerate=6`, `add=0`, `drop=0`, `error=1`
- unknown part: `docMetadata/LabelInfo.xml`
- content type: `application/vnd.ms-office.classificationlabels+xml`
- package issue refs: `Projekt 1!A1`
- the unknown part was copied through byte-for-byte and also surfaced as an explicit package graph error action.

Validation so far:

- `bun test fixtures/benchmarks/release-proof-index.test.ts`
- `bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`

## Confidence

Medium-high for owner-review evidence. The candidate proves Ascend can summarize a real unknown-part mutation boundary, but it does not close the fixture policy because the workbook is external, not vendored, and not approved as a tracked public binary fixture.

## Fold-in decision

Fold into release proof evidence only. This is not a new production surface and not approval to publish the stronger package-action claim. Handoff owners:

- product: decide whether to vendor or reject the external workbook as public fixture evidence.
- correctness: keep the unknown-part error semantics fail-closed unless owner policy accepts a broader fixture.
- release: keep non-attestation and non-trust wording.

## Next question

Can the owner-review evidence board identify the smallest remaining public fixture acquisition task that would close either safe-open `public-edge-fixtures` or package-action `edge-fixture-policy` without adding any new surfaces?
