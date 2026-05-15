# Layout Boolean Validation Boundary

## Question

Should visibility, row/column hide, and outline operations reject non-boolean public metadata before mutating workbook layout state?

## Hypothesis

Yes. These operations expose boolean public controls (`hidden`, `collapsed`, `summaryBelow`, `summaryRight`). Rejecting stringly or otherwise invalid values before mutation keeps the operation engine and journal exactness model aligned.

## External sources checked

- Microsoft Excel `Worksheet.Visible` property: https://learn.microsoft.com/office/vba/api/excel.worksheet.visible
- Microsoft Support on outlining/grouping worksheet data: https://support.microsoft.com/en-us/office/outline-group-data-in-a-worksheet-08ce98c4-0063-4d42-8ac7-8278c49e9aff
- Open XML SDK row outline-level attribute reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row.outlinelevel

## Why this matters to Ascend

Auditable mutation requires public operations to reject invalid metadata at the boundary and to explain non-representable journal cases consistently. Layout booleans are small, but accepting `"false"` as truthy would create surprising workbook mutations and weaken rollback evidence.

## Probe/implementation

An in-flight fix added:

- engine validation for invalid `hideSheet.hidden`, `hideRows.hidden`, `hideCols.hidden`, `groupRows.collapsed`, `groupRows.summaryBelow`, `groupCols.collapsed`, and `groupCols.summaryRight`.
- engine regression coverage proving invalid values reject without mutating sheet state, row/column definitions, outline properties, or sheet format properties.
- SDK journal exactness classification for the same invalid values as `UNSUPPORTED_VALUE`.

## Results

Focused validation:

- `bun test packages/engine/src/operations.test.ts -t "visibility and outline layout operations reject invalid booleans"`
- `bun test packages/sdk/src/journal-exactness.test.ts -t "classifies invalid row and column layout journals as unsupported values"`
- `bun test packages/sdk/src/journal-exactness.test.ts -t "classifies invalid sheet layout journals as unsupported values"`
- `bunx biome check packages/engine/src/operations/format-ops.ts packages/engine/src/operations/sheet-ops.ts packages/engine/src/operations.test.ts packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts`
- `bunx tsc --build`

Full-suite validation was also run while the diff was present:

- `bun run test:changed` passed with 5198 pass, 1 skip, 0 fail.

## Confidence

High. The fix is a narrow public-boundary validation change with no new product surface, and tests cover mutation prevention plus journal classification.

## Fold-in decision

Promote to correctness loop as auditable mutation hygiene. Do not promote new layout claims or protection/security wording.

## Next question

Can table-operation public metadata get the same validation/journal symmetry without touching unrelated release RC gate work?
