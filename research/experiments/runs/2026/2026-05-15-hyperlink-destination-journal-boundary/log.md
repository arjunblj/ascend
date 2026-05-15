# Hyperlink Destination Journal Boundary

## Question

Can Ascend honestly classify a `setHyperlink` journal as exact when the requested edit has neither an external URL nor an internal workbook location?

## Hypothesis

No. A hyperlink edit without a public destination is rejected by the engine, and OOXML hyperlink records express their destination through either a relationship target or a location attribute. The journal should therefore fail closed with `UNSUPPORTED_VALUE` instead of producing inverse operations that imply exact rollback.

## External sources checked

- Microsoft Open XML Spreadsheet `Hyperlink` class: https://learn.microsoft.com/th-th/dotnet/api/documentformat.openxml.spreadsheet.hyperlink?view=openxml-2.12.0
- OOXML relationships overview: https://ooxml.info/docs/9/9.2/
- OOXML hyperlink fundamentals: https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Hyperlinkshyperlink__topic_ID0ESHFO.html
- Microsoft Excel `HYPERLINK` function documentation: https://support.microsoft.com/en-gb/office/hyperlink-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f

## Why this matters to Ascend

Auditable mutation claims depend on not overstating inverse-journal exactness. Hyperlinks look simple, but a destinationless `setHyperlink` is not a reversible workbook edit through Ascend's public operation contract. Classifying it explicitly keeps the journal proof aligned with the engine's validation boundary.

## Probe/implementation

- Inspected `packages/engine/src/operations/format-ops.ts`; `setHyperlink` already rejects edits with no nonblank `url` or `location`.
- Updated `packages/sdk/src/journal.ts` so destinationless `setHyperlink` journal entries produce no inverse ops and carry an `UNSUPPORTED_VALUE` issue on the `hyperlinks` surface with reason `value-unsupported`.
- Changed the hyperlink exactness matrix from exact to conditional and added `value-unsupported` as the accepted loss reason.
- Added regression cases for omitted destination, blank URL, and blank location.

## Results

Validation:

```bash
bun test packages/sdk/src/journal-exactness.test.ts
bunx biome check packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts
bunx tsc --build
```

All passed. The new test verifies:

- `supported=true`
- `exact=false`
- `issueCount=1`
- `surfaces=['hyperlinks']`
- `reasons=['value-unsupported']`
- no exactness matrix violation

## Confidence

High for the local fix: engine validation, schema shape, and journal exactness now agree. Medium for broader hyperlink edge cases because imported hyperlink relationship quirks and relative targets were not expanded in this pass.

## Fold-in decision

Promote to the correctness loop as a small production fix. This is not a new user-facing surface; it only prevents a false exactness claim in journal evidence.

## Next question

Which remaining "exact" journal surfaces can still be made lossy by invalid public operation values before the engine rejects the edit?
