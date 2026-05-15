# Safe Open Public Fixture Replacement Scan

## Question

Can public binary fixtures replace the current code-generated signed and unknown package cases in the safe-open proof bundle?

## Hypothesis

Probably not from the current checked-in corpus. A local scan should prove whether any existing public XLSX/XLSM fixture already triggers the same open-plan risk families as the synthetic signed and unknown packages.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Aspose XLSX signature workflow, useful as a generated-signature contrast rather than a public fixture source: https://products.aspose.com/total/net/signature/xlsx/
- Public sample spreadsheet download pages such as Sample.Cat and SampleFile.com were checked, but the search did not surface a mature, provenance-clear signed XLSX fixture suitable for committing.

## Why this matters to Ascend

The safe-open release proof blocker is specific: signed and unknown-part cases are durable code-generated packages, not public binary fixtures. If existing public fixtures already cover those package features, the blocker can be retired. If not, release copy must keep disclosing synthetic edge proof.

## Probe/implementation

Folded in a small benchmark fixture-scan harness:

```bash
bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json
```

The harness walks checked-in `fixtures/xlsx` XLSX/XLSM files and runs `inspectWorkbookOpenPlan` on each one. It reports candidates whose risk families include `preservedSignature` or `preservedOther`.

Validation:

```bash
bun test fixtures/benchmarks/safe-open-fixture-scan.test.ts fixtures/benchmarks/safe-open-proof.test.ts fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/safe-open-fixture-scan.ts fixtures/benchmarks/safe-open-fixture-scan.test.ts
bunx tsc --build
```

## Results

Current scan:

- Scanned fixtures: 351.
- Rejected during scan: 2.
- Rejected fixtures: `fixtures/xlsx/calamine/pass_protected.xlsx`, `fixtures/xlsx/poi/protected_passtika.xlsx`.
- `preservedSignature` matches: 0.
- `preservedOther` matches: 0.
- Replacement status: `no-public-binary-replacement-found`.

This proves the current checked-in fixture corpus cannot replace the synthetic signed and unknown safe-open cases today.

## Confidence

High for the checked-in fixture corpus. Medium for the wider public internet because search results found generic sample files and signature-generation services, not a clearly licensed public signed workbook fixture with stable provenance.

## Fold-in decision

Promote to product/performance proof packaging. Keep the synthetic blocker disclosed. Do not download or commit random signed workbooks without clear provenance and license.

## Next question

Can a controlled public fixture policy define when generated edge packages are acceptable release proof versus when binary public fixtures are required?
