# Full Scalar Formula Scan Hoist

## Question

Can the full-scalar worksheet byte parser avoid repeated per-row formula scans on formula-free sheets while preserving formula fallback behavior?

## Hypothesis

Yes. SpreadsheetML stores formulas in `<f>` elements under cells. If a worksheet `sheetData` region contains no `<f>` elements, the full-scalar parser can skip row-local formula checks and still fall back to the XML path when formulas exist.

## External sources checked

- Microsoft Open XML formulas documentation: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- Microsoft Open XML SDK `CellFormula`: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula
- OOXML formula element reference: https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_f_topic_ID0E6TY4.html

## Why this matters to Ascend

Ascend's performance North Star depends on real XLSX read paths staying fast without hiding workbook semantics. Formula-free dense value sheets are common; avoiding repeated formula scans improves the no-formula fast path while preserving the fallback boundary for formula-bearing worksheets.

## Probe/implementation

Implemented in `packages/io-xlsx/src/reader/sheet.ts`:

- locate the `</sheetData>` close once in `parseSheetFullScalarBytes`.
- scan the `sheetData` body once for `<f>` and return `null` to the full XML path when formulas are present.
- pass `formulasKnownAbsent=true` into row parsing on the full-scalar path.
- bound row-open scanning to the sheetData close and reuse the precomputed close for `maxRows`.

## Results

Focused validation:

- `bun test packages/io-xlsx/src/reader/reader.test.ts -t "formula-only byte parser hydrates no-formula sheets and falls back for formulas"`
- `bun test packages/io-xlsx/src/reader/reader.test.ts -t "values mode reads dense workbooks successfully across repeated runs"`
- `bun test packages/io-xlsx/src/reader/reader.test.ts -t "values mode reads sequential A1-style refs beyond Z without shifting columns"`
- `bunx biome check packages/io-xlsx/src/reader/sheet.ts packages/io-xlsx/src/reader/reader.test.ts`
- `bunx tsc --build`

Local diagnostic probe:

```bash
bun run fixtures/benchmarks/xlsx-read-phase.ts --rows 5000 --cols 20 --workload dense-values --read-source raw-ooxml --phase read --repeat 3 --warmup 1 --validation-mode sample --json
```

Result summary:

- `readXlsxMedianMs=8.356042`
- `readXlsxCellsPerSecondMedian=11967388.387947304`
- input bytes `281347`

## Confidence

Medium. Correctness confidence is high because formula fallback tests pass; performance confidence is diagnostic-only because there is no before/after baseline in this log.

## Fold-in decision

Promote to performance loop as a narrow IO fast-path improvement. Do not turn this into a release performance claim until an approved before/after benchmark or release-environment run exists.

## Next question

Should the RC gate script become the next release-proof fold-in, or should it stay uncommitted until product/release owners approve publication-policy wording?
