# Metadata Duplicate Journal Boundary

Date: 2026-05-15

## Question

Can mutation journals distinguish non-suffix metadata order loss from duplicate metadata that public operations cannot target exactly?

## Hypothesis

Yes. Duplicate data validations or conditional-format blocks with the same public range should be classified as `metadata-duplicate`, not merely `metadata-order`, because the public delete/restore operation cannot select which duplicate metadata entry is being restored.

## External sources checked

- Microsoft Open XML `DataValidation` documents `sqref` as the sequence of references affected by a data-validation element: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidation?view=openxml-2.20.0
- Microsoft Open XML conditional formatting documentation describes conditional-formatting metadata as worksheet XML elements over cell/range collections: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting

## Why this matters to Ascend

The auditable package-part mutation claim depends on precise loss reasons. Order loss and duplicate-selector loss are different boundaries: order loss says content can be restored at the wrong list position, while duplicate loss says the public selector cannot identify one metadata entry among multiple equal-range entries.

## Probe/implementation

Folded in a scoped SDK journal classification fix:

- `dataValidationRestoreOrderIssues` now reports `reason: "metadata-duplicate"` when the preimage range matches duplicate worksheet validation metadata.
- `conditionalFormatRestoreOrderIssues` now reports `reason: "metadata-duplicate"` when the preimage range itself is duplicated.
- Added journal exactness tests for duplicate data-validation deletion and duplicate conditional-format deletion.

## Results

Validation passed:

```bash
bun test packages/sdk/src/journal-exactness.test.ts -t "duplicate|generated lossy journals"
bunx biome check packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts
bunx tsc --build packages/sdk/tsconfig.json
```

Observed proof:

- Duplicate `Sheet1!A1:A1` data validations now produce `LOSSY_INVERSE`, `surface: "data-validations"`, `reason: "metadata-duplicate"`.
- Duplicate `Sheet1!A1:A1` conditional formats now produce `LOSSY_INVERSE`, `surface: "conditional-formats"`, `reason: "metadata-duplicate"`.

## Confidence

High for duplicate same-range metadata selectors. Medium for richer duplicate definitions involving same range but different rule priority or x14 metadata; those should remain explicit future probes.

## Fold-in decision

Promote to correctness loop and commit. This is a precision fix to the existing mutation journal vocabulary, not a new surface.

## Next question

Should property-based journal-law generators include duplicate metadata cases for validations and conditional formats before they are promoted as a correctness proof?
