# Sheet Protection Validation Boundary

## Question

Should Ascend reject invalid `setSheetProtection` public metadata before mutating sheet protection state, and should journal exactness classify those same invalid values?

## Hypothesis

Yes. Sheet protection options are boolean public metadata, and passwords should enter Ascend's public operation layer as strings. Rejecting invalid values before mutation keeps the operation engine and mutation journal aligned with the auditable mutation claim.

## External sources checked

- Microsoft Excel `Worksheet.Protect` API, including password and allow-sort/filter/format options: https://learn.microsoft.com/office/vba/api/Excel.Worksheet.Protect
- OOXML `sheetProtection` element reference: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_sheetProtection_topic_ID0ENIH5.html
- Microsoft Open XML SDK `sheetProtection` element listing: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.linq.x.sheetprotection

## Why this matters to Ascend

Ascend's release claim for auditable mutation depends on public operations rejecting non-representable metadata before mutation and explaining journal limits consistently. Sheet protection sits on a user-visible workbook boundary, so accepting malformed public metadata would weaken the correctness story even if XLSX writing later failed.

## Probe/implementation

An in-flight fix added:

- `validateSheetProtectionInput` in `packages/engine/src/operations/sheet-ops.ts`.
- engine regression coverage for invalid password, invalid options container, and non-boolean option values.
- SDK journal exactness classification for the same invalid public inputs as `UNSUPPORTED_VALUE` on `sheet-layout`.

## Results

Focused validation:

- `bun test packages/engine/src/operations.test.ts -t "setSheetProtection rejects invalid"`
- `bun test packages/sdk/src/journal-exactness.test.ts -t "unsupported value"`
- `bunx biome check packages/engine/src/operations/sheet-ops.ts packages/engine/src/operations.test.ts packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts`

Full-suite validation was already run while this diff was present:

- `bun run test:changed` passed with 5196 pass, 1 skip, 0 fail.

## Confidence

High. The fix is small, the option field list matches Ascend's public `SheetProtectionOptions`, and the tests assert no mutation occurs after invalid inputs.

## Fold-in decision

Promote to correctness loop. This is a narrow validation fix under the auditable mutation claim, not a new product surface and not stronger protection/security wording.

## Next question

Can the remaining metadata operation validators be checked for the same engine/journal symmetry without adding any release claim surface?
