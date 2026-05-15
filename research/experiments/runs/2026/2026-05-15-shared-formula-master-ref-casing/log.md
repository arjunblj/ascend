# Shared formula master ref casing

## Question

Do engine mutation and SDK journal preimage paths materialize imported shared-formula groups when member `masterRef` values are sheet-qualified with different casing from the workbook sheet name?

## Hypothesis

Yes. The same case-insensitive sheet-name semantics used by verifier binding checks should also protect mutation materialization and journal evidence. A member with `masterRef: "sheet1!$A$1"` should resolve to a master stored as `Sheet1!A1`, detach both cells before a literal replacement, and expose both preimages in the journal.

## External sources checked

- Microsoft Open XML formulas overview describes SpreadsheetML formulas, cell references, and formulas stored in worksheet XML: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- Open XML SDK `CellFormula` docs describe the `shared` formula type, `ref`, and shared-index attributes: https://learn.microsoft.com/es-es/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1
- Apache POI formula type docs describe shared formulas as a file-size optimization stored as shared formula records: https://poi.apache.org/apidocs/3.17/org/apache/poi/ss/formula/FormulaType.html

## Why this matters to Ascend

Auditable package-part mutation and safe formula metadata handling both depend on detaching imported binding groups before public operations rewrite or replace cells. If casing drift in sheet-qualified `masterRef` values splits a shared group, Ascend could produce incomplete preimages or leave stale formula metadata behind.

## Probe/implementation

Inspected current shared-formula materialization helpers and the dirty in-flight tests in:

- `packages/engine/src/operations.test.ts`
- `packages/sdk/src/interactive-contract.test.ts`

The production helper implementation was already present in the current tree; this checkpoint adds regression proof only:

- Engine `setCells` case: master `Sheet1!A1`, member `sheet1!$A$1`, literal replacement at `A2`, expected affected cells `A1` and `A2`, and formula metadata detached from both.
- SDK journal case: same imported metadata, journal-enabled `setCells`, expected preimage cells `A1` and `A2`, shared-formula lossy issues for both refs, and detached formula metadata.

Validation:

```bash
bun test packages/engine/src/operations.test.ts -t "case-insensitive sheet-qualified master refs"
bun test packages/sdk/src/interactive-contract.test.ts -t "case-insensitive sheet-qualified master refs"
```

## Results

- Engine targeted test passed: 1 test, 7 assertions.
- SDK targeted test passed: 1 test, 9 assertions.
- The proof is test-only and does not add a public surface.

## Confidence

High for the `setCells` shared-formula materialization and SDK journal preimage paths. Medium for the entire operation set because this probe does not exercise every operation that can detach imported shared formulas; it covers the most direct literal-replacement path and journal evidence path.

## Fold-in decision

Promote to correctness loop as regression coverage. This complements the verifier case-insensitive binding-ref fix without expanding formula rename or package-action release claims.

## Next question

Should the correctness loop add a small matrix for case-insensitive sheet-qualified binding refs across `setFormula`, `fillFormula`, `copyRange`, and `renameSheet`, or is the shared helper coverage plus this direct `setCells` regression sufficient?
