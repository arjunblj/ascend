# Axis Layout Validation Boundary

## Question

Can row/column layout operations reject invalid coordinates and dimensions before mutating workbook metadata?

## Hypothesis

Yes. The engine can fail closed on non-integer or negative row/column indexes, zero or invalid spans, and negative or non-finite row/column dimensions before touching cells, row metadata, column metadata, or outline metadata.

## External sources checked

- Microsoft Support, change column width and row height: https://support.microsoft.com/en-au/office/change-the-column-width-and-row-height-72f5e3cc-994d-43e8-ae58-9774a0905f46
- Microsoft Learn Open XML `Row.Height` property: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row.height?view=openxml-3.0.1
- Microsoft Learn Office Standards note for Open XML column width and formatting: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/302056b0-09cb-49bb-b8fe-6e7ed0ef3f8d
- ECMA-376 Office Open XML standards landing page: https://ecma-international.org/publications-and-standards/standards/ecma-376/

## Why this matters to Ascend

Auditable mutation claims depend on rejected operations being boring: an invalid layout edit should not partially dirty workbook metadata and then rely on later rollback or package evidence to explain the damage. This is correctness hygiene under the auditable package-part mutation claim, not a new product surface.

## Probe/implementation

- Inspected the in-flight engine patch touching:
  - `packages/engine/src/operations/structural-ops.ts`
  - `packages/engine/src/operations/sheet-ops.ts`
  - `packages/engine/src/operations/format-ops.ts`
  - `packages/engine/src/operations.test.ts`
  - `packages/sdk/src/journal.ts`
  - `packages/sdk/src/journal-exactness.test.ts`
- Finished the patch by validating:
  - structural row/column shifts: integer, non-negative `at`, positive integer `count`;
  - hide row/column spans: integer, non-negative `at`, positive integer `count`;
  - row heights and column widths: integer, non-negative index and non-negative finite dimension;
  - outline groups: integer row/column endpoints before writing outline metadata.
- Kept the behavior scoped to validation errors and added focused operation-test coverage for metadata not being written on rejection.
- Folded the same boundary into SDK journal analysis so invalid layout operations are reported as supported-but-inexact `UNSUPPORTED_VALUE` journal issues on `row-layout` or `column-layout`, with `value-unsupported` as the reason.

## Results

- `bun test packages/engine/src/operations.test.ts`: 288 pass, 0 fail.
- `bun test packages/engine/src/operations.test.ts packages/sdk/src/journal-exactness.test.ts`: 353 pass, 0 fail.
- `bunx biome check packages/engine/src/operations.test.ts packages/engine/src/operations/format-ops.ts packages/engine/src/operations/sheet-ops.ts packages/engine/src/operations/structural-ops.ts packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts`: passed.
- `bunx tsc --build`: passed.
- `bun run test:changed`: rerun passed with 5179 pass, 1 skip, 0 fail.
- The first local test attempt failed because a redundant added test tried to call a non-existent `SparseGrid.entries()` helper. I removed that duplicate and relied on the broader row/column layout rejection test.
- The first `test:changed` attempt reported stale post-write failures under Bun's `--only-failures`; targeted repros passed and the immediate rerun passed the full affected suite.

## Confidence

Medium-high for the targeted validation boundary. The implementation is small and covered by engine operation tests, but broader validation consistency across every workbook operation is still a separate correctness program.

## Fold-in decision

Promote to correctness loop. Commit as a tiny engine validation fix under the auditable mutation claim. Do not promote a new release claim or user-facing surface.

## Next question

Can the claim board be refreshed from the current machine proof so stale human text does not contradict `release-proof-index` on package-action streaming coverage?
