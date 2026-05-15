# Safe Unknown Workbook Opening Proof Benchmark

## Question

Can Ascend's safe unknown workbook opening claim be supported with public-fixture latency evidence without adding another production surface?

## Hypothesis

Yes. `inspectWorkbookOpenPlan()` already works from package-level structure before workbook hydration. A local public-fixture probe should show that it is materially cheaper than full hydration while still routing active-content workbooks to metadata-only review.

## External sources checked

- SheetJS parse options: https://docs.sheetjs.com/docs/api/parse-options
- SheetJS workbook object and VBA metadata behavior: https://docs.sheetjs.com/docs/csf/book/
- openpyxl tutorial warning for unsupported item loss and `keep_vba`: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- openpyxl optimized read modes: https://openpyxl.pages.heptapod.net/openpyxl/optimized.html
- ExcelJS read/streaming/ignoreNodes positioning: https://www.npmjs.com/package/exceljs
- LibreOffice Calc guide for Excel file open/edit/save breadth: https://documentation.libreoffice.org/assets/Uploads/Documentation/en/CG7.0/CG70-CalcGuide.pdf

## Why this matters to Ascend

The claim ladder ranks "safe unknown workbook opening" as Ascend's most credible product-shaped claim today. The claim is stronger if it has both safety evidence and cost evidence: agents can ask a package-level routing question before reading workbook cells or planning edits.

## Probe/implementation

- Inspected `packages/sdk/src/open-plan.ts` and existing CLI/API/SDK open-plan tests.
- Ran an inline Bun probe over public fixtures:
  - `fixtures/xlsx/poi/SampleSS.xlsx`
  - `fixtures/xlsx/calamine/vba.xlsm`
  - `fixtures/xlsx/poi/ExcelPivotTableSample.xlsx`
  - `fixtures/xlsx/poi/WithChart.xlsx`
  - `fixtures/xlsx/poi/WithTable.xlsx`
- Each fixture was warmed once, then measured over 9 iterations:
  - `inspectWorkbookOpenPlan(bytes, { intent: "edit-plan" })`
  - `AscendWorkbook.open(bytes, { mode: "full" })`
- No production changes were made.

## Results

Local probe output:

| Fixture | Recommended mode | Review before hydration | Parts | Relationships | Median open-plan ms | Median full-open ms | Full/open-plan ratio |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| clean `SampleSS.xlsx` | formula | false | 13 | 10 | 0.215 | 2.410 | 11.22x |
| macro `vba.xlsm` | metadata-only | true | 12 | 9 | 0.115 | 1.595 | 13.83x |
| pivot `ExcelPivotTableSample.xlsx` | formula | false | 27 | 19 | 0.187 | 2.467 | 13.20x |
| chart `WithChart.xlsx` | formula | false | 15 | 10 | 0.095 | 1.286 | 13.47x |
| table `WithTable.xlsx` | formula | false | 15 | 11 | 0.096 | 1.208 | 12.61x |

Focused validation passed:

```bash
bun test packages/sdk/src/open-plan.test.ts
bun test packages/sdk/src/open-plan.test.ts apps/cli/src/cli.test.ts -t "open-plan" apps/api/api.test.ts -t "open-plan"
```

The existing tests prove SDK, CLI, and API open-plan behavior, including macro routing to metadata-only review.

## Confidence

Medium. The fixture sample is public and repeatable, and the timing gap is large enough to justify the claim direction. It is still a local microbenchmark, not a formal benchmark threshold, and it does not include a malformed workbook case yet.

## Fold-in decision

Keep as research evidence and hand off to the product/performance loop. Do not add another production surface. The next implementation should be a fixture-backed proof bundle/report over existing `open-plan` surfaces if product wants to promote this claim.

## Next question

Should the next product loop package token-bounded agent view and retained viewport patch history as documented agent contracts, or should correctness continue deepening package action proof with per-part digests?
