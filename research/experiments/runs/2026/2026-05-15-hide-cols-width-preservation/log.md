# Hide Columns Width Preservation

Date: 2026-05-15

## Question

Can Ascend preserve a column's public width metadata when `hideCols` creates or splits a `<col>` definition for a column that was previously resized through `setColWidth`?

## Hypothesis

Yes. If the sparse public `sheet.colWidths` map already records a width for a column, `hideCols` should carry that width and `customWidth` flag into the generated column definition instead of creating a hidden-only definition that forgets the user-visible width.

## External Sources Checked

- Microsoft Open XML `Column` documentation describes `<col>` as the worksheet column width/formatting record and lists `customWidth`, `hidden`, `min`, `max`, and `width` attributes: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-3.0.1
- SheetJS column properties documentation treats width and hidden column metadata as separate column properties that can coexist in formats that support them: https://docs.sheetjs.com/docs/csf/features/colprops/

## Why This Matters To Ascend

This supports the "auditable package-part mutation" and correctness loops at a small, concrete level: mutation operations should not silently drop user-visible layout metadata while adding another layout flag. Preserving width through hide/unhide also makes inverse journal and package write claims more defensible for real workbook layout edits.

## Probe/Implementation

Folded in a narrow engine fix:

- `setColHidden` now copies public column width metadata from `sheet.colWidths` when it creates a new hidden column definition.
- When splitting an existing column definition that lacks width, it also fills width/customWidth from the public width map for the target column.
- Added an engine regression test proving `setColWidth` followed by hide/unhide preserves `sheet.colWidths` and `sheet.colDefs` width metadata.

Dirty but not folded in this checkpoint:

- A broader journal-exactness test expansion is present locally, but one new conditional-format package-byte exactness case currently fails. That evidence should become its own correctness loop instead of being smuggled into this width-preservation fix.

## Results

Validation passed:

```bash
bun test packages/engine/src/operations.test.ts -t "hideCols preserves public column widths|setColWidth updates imported column definition"
bunx biome check packages/engine/src/operations/sheet-ops.ts packages/engine/src/operations.test.ts
bunx tsc --build packages/engine/tsconfig.json
```

Observed behavior:

- After `setColWidth(Sheet1, B, 12)`, `hideCols(B)` leaves `sheet.colWidths.get(1) === 12`.
- The generated column definition is `{ min: 1, max: 1, width: 12, customWidth: true, hidden: true }`.
- Unhiding leaves `{ min: 1, max: 1, width: 12, customWidth: true }`.

## Confidence

High for the narrow engine behavior. Medium for broader package-byte claims because the local journal-exactness probe found a separate conditional-format saved-byte restoration gap.

## Fold-In Decision

Promote to correctness loop and commit the scoped engine fix plus regression test. Do not promote the broader journal-exactness additions until the failing conditional-format exactness case is either fixed or explicitly classified as lossy.

## Next Question

Can conditional-format replacement/deletion journals restore saved package bytes exactly, or should that operation family be classified with an honest lossy/package-byte boundary?
