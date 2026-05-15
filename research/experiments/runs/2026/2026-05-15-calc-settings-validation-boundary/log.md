# Calc Settings Validation Boundary

## Question

Should the in-flight workbook calculation-settings validation patch be folded into the auditable mutation claim?

## Hypothesis

Yes. Calculation settings are workbook metadata, and invalid public inputs should fail before mutation while still producing journal issue evidence with the shared unsupported-value vocabulary.

## External sources checked

- Microsoft Open XML SDK `CalculationProperties`, which maps workbook `calcPr` attributes including `calcMode`, `calcCompleted`, `calcOnSave`, `forceFullCalc`, `fullCalcOnLoad`, and iterative calculation fields: https://learn.microsoft.com/ru-ru/dotnet/api/documentformat.openxml.spreadsheet.calculationproperties?view=openxml-3.0.1
- Microsoft Open XML SDK `FullCalculationOnLoad`, which exposes `fullCalcOnLoad` as a BooleanValue schema attribute: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetcalculationproperties.fullcalculationonload?view=openxml-3.0.1
- OOXML `calcPr` schema summary for enum, unsigned integer, and boolean attribute shapes: https://www.datypic.com/sc/ooxml/e-ssml_calcPr-1.html

## Why this matters to Ascend

Auditable package-part mutation depends on rejecting invalid workbook metadata before it can produce impossible workbook state. Calc settings affect formula recalc behavior and are visible in workbook XML, so accepting invalid enum, boolean, or iterative values would weaken the "plan/commit/journal is explainable" claim.

## Probe/implementation

- Inspected the dirty in-flight patch in `packages/engine/src/operations/workbook-ops.ts`.
- Finished the boundary by adding SDK journal unsupported-value classification for invalid `calcMode`, boolean flags, `dateSystem`, malformed `iterativeCalc`, bad `iterativeCalc.enabled`, and non-finite `iterativeCalc.maxChange`.
- Added engine tests proving invalid calc settings reject before mutation.
- Added journal exactness cases proving rejected calc-setting inputs remain classified as workbook-metadata `UNSUPPORTED_VALUE` issues.

## Results

- Targeted validation passed:
  - `bun test packages/engine/src/workbook-ops.test.ts packages/sdk/src/journal-exactness.test.ts`
  - `bunx biome check packages/engine/src/operations/workbook-ops.ts packages/engine/src/workbook-ops.test.ts packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts`
- The patch is correctness hygiene under auditable mutation. It does not add a new product surface and does not change benchmark thresholds.
- The separate dirty `packages/io-xlsx/src/reader/sheet.ts` performance edit was left unstaged because it is a different in-flight change.

## Confidence

High for the validation and journal boundary. The probe is local and targeted, and the external schema references support the enum/boolean/numeric shape being enforced.

## Fold-in decision

Promote to correctness loop and commit as a small production fix. This supports auditable workbook metadata mutation, not a new release claim.

## Next question

Can the safe-open external unknown-part workbook candidate become machine-readable owner-review evidence in the release proof index without satisfying the public-edge fixture gate?
