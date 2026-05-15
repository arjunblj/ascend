# Style Edit Formula Binding Preservation

## Question

Do style-only operations preserve representative imported formula binding metadata?

## Hypothesis

Yes. Style changes should not detach or rewrite formula-binding metadata for shared formulas, dynamic arrays, blocked spills, data tables, or legacy arrays. If they do, the auditable mutation claim would blur harmless presentation edits with semantic formula edits.

## External sources checked

- Microsoft Open XML formula guidance, which documents SpreadsheetML cell formula markup and formula metadata context: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- Microsoft Open XML SDK `CellFormula`, which maps formula references and array/shared formula attributes to workbook XML: https://learn.microsoft.com/es-es/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1
- OOXML formula type summary for shared, array, and data-table formula classes: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_CellFormulaType_topic_ID0ETSDFB.html

## Why this matters to Ascend

Formula binding metadata is part of Ascend's correctness moat. Style-only edits should be safe presentation changes; if they disturb formula binding metadata, journals and package proof would need to classify a style edit as semantically lossy.

## Probe/implementation

- Inspected the in-flight test-only patch in `packages/engine/src/operations.test.ts`.
- Added a regression matrix for `setNumberFormat` and `setStyle` over:
  - shared formula member;
  - dynamic array spill member;
  - blocked spill anchor;
  - data-table metadata;
  - legacy array metadata.
- The test records `formulaInfo` before the style operation and asserts the same metadata remains after the style operation.

## Results

- Targeted validation passed:
  - `bun test packages/engine/src/operations.test.ts -t "style setters preserve representative formula binding metadata"`
  - `bunx biome check packages/engine/src/operations.test.ts`
- The proof landed as `f170a930 test(engine): prove style edits preserve formula bindings`.
- Fold-in scope is correctness test coverage only. No production behavior, public surface, or benchmark threshold changed.

## Confidence

Medium-high. The matrix covers the representative local formula-binding classes already used by Ascend tests, but it is still synthetic test coverage rather than a broad imported workbook corpus.

## Fold-in decision

Promote to correctness loop as regression coverage under auditable mutation. Do not promote a new formula-intelligence claim.

## Next question

Should package-action `unknown-part-error` get an external-candidate owner-review record based on the ExcelForge local mutation probe, while still keeping the edge-fixture policy gate missing?
