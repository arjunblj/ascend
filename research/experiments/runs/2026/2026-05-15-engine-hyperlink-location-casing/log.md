# Engine Hyperlink Location Casing

## Question

Do engine sheet-topology operations rewrite internal hyperlink locations case-insensitively when the hyperlink target uses a different sheet-name casing?

## Hypothesis

They should. Hyperlink locations such as `sheet1!A2` represent the same worksheet as `Sheet1!A2`; insert and rename operations should not leave those links stale.

## External sources checked

- Microsoft documents Excel links to a place in the current workbook as worksheet plus cell-reference targets: https://support.microsoft.com/en-gb/office/work-with-links-in-excel-7fc80d8d-68f9-482f-ab01-584c44d72b3e
- Microsoft worksheet references define references as sheet/cell targets, supporting symbolic interpretation rather than raw text equality: https://learn.microsoft.com/en-us/office/client-developer/excel/worksheet-references
- Microsoft Open XML Hyperlink API documents spreadsheet hyperlink `ref` and `location` metadata as workbook hyperlink state: https://learn.microsoft.com/th-th/dotnet/api/documentformat.openxml.spreadsheet.hyperlink?view=openxml-2.12.0

## Why this matters to Ascend

Safe workbook mutation is not just formulas. Internal hyperlinks are user-facing navigation metadata, and stale targets after row/column inserts or sheet rename weaken the preservation-first story. This is still a tiny correctness fix, not a new product claim.

## Probe/implementation

Commit `354aa1a0 fix(engine): match hyperlink sheet refs case-insensitively` updated `packages/engine/src/structural/sheet-topology.ts`:

- `renameHyperlinkLocation()` compares hyperlink target sheet names case-insensitively.
- `shiftHyperlinkLocation()` compares hyperlink target sheet names case-insensitively.
- The comparison preserves the existing target sheet casing when shifting and uses the new sheet name when renaming.

Regression probes in `packages/engine/src/operations.test.ts` changed internal hyperlink targets to lowercase sheet qualifiers:

- `sheet1!A2` shifts to `sheet1!A4` after row insert.
- `sheet1!A1` renames to `Data!A1` after sheet rename.

## Results

Focused validation passed:

```bash
bun test packages/engine/src/operations.test.ts -t "hyperlink|renameSheet updates validation"
```

Result: 11 passed, 0 failed.

## Confidence

High for hyperlink location shift and rename paths covered by existing engine operation tests. Medium for every metadata surface until the remaining sheet-reference comparators are audited.

## Fold-in decision

Folded into the correctness loop as a tiny fix. It supports preservation-first mutation and package audit credibility, but it is not an implementation-loop handoff or a new surface.

## Next question

After this in-flight fix is committed, can the release claim board stay pinned to the top two owner loops and explicitly avoid promoting formula rename or practical latency into release wording?
