# Data Validation Order Journal Boundary

Date: 2026-05-15

## Question

Can data-validation replacement/deletion journals honestly distinguish exact inverse restoration from cases where public operations can restore validation content but not saved package order?

## Hypothesis

Yes. Classic worksheet data validations are serialized as ordered SpreadsheetML metadata. Ascend should keep suffix replacement/deletion exact, but classify non-suffix replacement/deletion as `LOSSY_INVERSE` with `surface: "data-validations"` and `reason: "metadata-order"` because public inverse operations append restored validations.

## External Sources Checked

- Microsoft Open XML SDK `DataValidation` maps to the SpreadsheetML `dataValidation` element and documents `sqref` as the sequence of references affected by the validation: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidation?view=openxml-2.20.0
- Microsoft Open XML SDK `DataValidations` is the worksheet container for child `DataValidation` elements: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidations?view=openxml-3.0.1

## Why This Matters To Ascend

This is a claim-stewardship fix for auditable package-part mutation. A journal that claims exactness must not silently turn `[A validation, B validation]` into `[B validation, A validation]` after inverse apply just because the public operation surface can only re-add the deleted or replaced validation at the tail.

## Probe/Implementation

Folded in a scoped SDK journal change:

- `journalSetDataValidation` now adds a metadata-order lossy issue when a replaced validation is not a suffix of the current validation list.
- `journalDeleteDataValidation` now adds the same issue for non-suffix deletes.
- The representative exact saved-byte test was narrowed to suffix replacement/deletion, which public inverse operations can restore byte-for-byte.
- Added direct classification coverage for non-suffix data-validation delete and replacement using the same two-validation workbook.

## Results

Validation passed:

```bash
bun test packages/sdk/src/journal-exactness.test.ts
bun test packages/sdk/src/interactive-contract.test.ts -t "journal inverse ops restore data validation changes|data validation journals mark absent default attributes lossy"
bunx biome check packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts
bunx tsc --build packages/sdk/tsconfig.json
```

Observed proof:

- Non-suffix deletion of `Sheet1!A1:A1` while `Sheet1!B1:B1` follows now yields `LOSSY_INVERSE`, `surface: "data-validations"`, `reason: "metadata-order"`, `refs: ["Sheet1!A1:A1"]`.
- Non-suffix replacement of `Sheet1!A1:A1` while `Sheet1!B1:B1` follows yields the same metadata-order issue.
- Suffix replacement/deletion remains in the representative exact saved-byte test and passes byte restoration.

## Confidence

High for classic worksheet data-validation ordering. Medium for extended x14 validation variants, which remain intentionally covered by separate x14 loss reasons rather than this classic metadata-order proof.

## Fold-In Decision

Promote to correctness loop and commit. This is not a new surface; it is a narrower honesty fix for the existing auditable mutation journal.

## Next Question

Should the next claim-steward block package this data-validation and conditional-format ordering boundary into the release-claim board, or should property-based journal-law generators prove these ordering boundaries across generated validation lists first?
