# Workbook Property Package State Journal

## Question

Do saved-source mutation journals correctly mark workbook property edits as package-state lossy evidence instead of presenting the public inverse as fully package-exact?

## Hypothesis

No. `setWorkbookProperties` mutates workbook-level package metadata such as the `workbookPr` `date1904` flag, so a saved-source journal can have a useful inverse operation while still needing a `package-part-preservation` issue tied to workbook properties.

## External sources checked

- Microsoft Open XML SDK `WorkbookProperties.Date1904` documents `date1904` as the schema attribute on workbook properties: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookproperties.date1904?view=openxml-3.0.1
- Microsoft `[MS-XLS] Date1904` records define the 1900 and 1904 workbook date systems and their serial bases: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/4a5e900a-0eb0-4355-8fc1-81aab8f46e8b
- SheetJS write options document workbook properties and date-code handling as file-level/write concerns: https://docs.sheetjs.com/docs/api/write-options/
- SheetJS date docs describe XLSX/XLSM dates as numeric date systems with workbook-level 1900/1904 interpretation: https://docs.sheetjs.com/docs/csf/features/dates/

## Why this matters to Ascend

The auditable package-part mutation claim depends on honest boundaries. A workbook property inverse can restore Ascend's in-memory model, but saved-source package bytes may still differ because workbook metadata lives in package XML. The journal must surface that as package-state evidence for agents and users.

## Probe/implementation

- Inspected the saved-source package-state journal path in `packages/sdk/src/journal.ts`.
- Confirmed `setWorkbookProperties` was missing from `savedSourcePackageStateRefsForOp`.
- Folded in the narrow production fix in commit `efd95a7b fix(sdk): locate workbook property package journal issues`.
- Added a saved-source exactness case for `setWorkbookProperties` expecting `refs: ['workbook:properties']`.

## Results

The fix maps `setWorkbookProperties` to `workbook:properties`, so saved-source journals now classify workbook property edits as package-part preservation risks when inverse operations cannot restore saved package bytes exactly.

Validation:

```bash
bun test packages/sdk/src/journal-exactness.test.ts -t "saved-source package state journals|workbook properties|representative exact saved journals"
bunx biome check packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts
bunx tsc --build
bun run test:changed
```

Observed validation result: targeted journal tests passed, Biome passed, typecheck passed, and `test:changed` ran the full suite with `5017 pass`, `1 skip`, `0 fail`.

## Confidence

High. The patch is a single operation-to-ref classification with direct saved-source byte exactness coverage.

## Fold-in decision

Promote to correctness loop and keep folded into production. This strengthens the second-ranked release claim, "auditable package-part mutation," without adding a new SDK/CLI/API/MCP surface.

## Next question

Can the auditable package-part proof report add one compatibility check that correlates package action evidence with journal package-state issues without expanding the public surface?
