# Binding range casing detach

## Question

Do formula-binding detachment helpers handle sheet-qualified binding ranges case-insensitively for blocked spills and data tables in both engine mutations and SDK journals?

## Hypothesis

Yes. Imported metadata can carry `ref` or `blockingRefs` as `sheet1!A1:A3` while the workbook sheet is `Sheet1`. Detachment should still discover the binding group, clear stale metadata before public edits, and report journal preimages/issues for the binding owner cells.

## External sources checked

- Microsoft Open XML formulas overview describes worksheet XML formulas, cell references, and names: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- Open XML SDK `CellFormula` docs describe formula types, reference ranges, shared indexes, and data-table formula metadata: https://learn.microsoft.com/es-es/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1
- Apache POI formula type docs describe shared formulas as file-size optimized formula metadata, reinforcing that imported formula binding metadata needs special handling: https://poi.apache.org/apidocs/3.17/org/apache/poi/ss/formula/FormulaType.html
- Microsoft worksheet rename rules document user-visible worksheet-name constraints, but not a separate case-sensitive sheet identity model: https://support.microsoft.com/en-gb/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9

## Why this matters to Ascend

Ascend's mutation claim depends on not silently leaving stale formula-binding metadata behind. Shared formulas, dynamic spills, blocked spills, and data tables are imported workbook semantics that public operations cannot always restore exactly. The correct behavior is to detach affected binding groups and record lossy journal evidence, even when producer casing differs.

## Probe/implementation

Production helpers updated:

- `packages/engine/src/operations/helpers.ts`
- `packages/sdk/src/journal.ts`

Committed production change:

```text
a5a970bb fix(engine): match binding ranges case-insensitively
```

Both `formulaBindingRangeContainsCell` helpers now compare parsed sheet-qualified range sheets through the existing case-insensitive sheet-name comparison instead of strict string equality.

Regression coverage:

- `packages/engine/src/operations.test.ts`
  - blocked-spill `ref`/`blockingRefs` as `sheet1!...`;
  - data-table `ref` as `sheet1!C3:C5`.
- `packages/sdk/src/interactive-contract.test.ts`
  - journal preimages and lossy issues for data-table and blocked-spill metadata with case-insensitive sheet-qualified ranges.

Validation:

```bash
bun test packages/engine/src/operations.test.ts -t "case-insensitive sheet-qualified ranges"
bun test packages/sdk/src/interactive-contract.test.ts -t "lossy imported formula-binding metadata preimages|data-table and blocked-spill preimages"
bunx biome check packages/engine/src/operations.test.ts packages/engine/src/operations/helpers.ts packages/sdk/src/interactive-contract.test.ts packages/sdk/src/journal.ts
bunx tsc --build
bun run test:changed
```

## Results

- Engine targeted tests passed: 2 tests, 11 assertions.
- SDK targeted tests passed: 2 tests, 51 assertions.
- Biome passed for all four touched production/test files.
- TypeScript build passed.
- `bun run test:changed` passed: 3988 pass, 1 skip, 0 fail across 136 files.
- The change is internal correctness behavior only; no API, CLI, MCP, or formula rename surface was added.

## Confidence

High for `setCells` blocked-spill and data-table detachment plus SDK journal evidence. Medium for every operation path because the probe does not enumerate all mutation operations, but both modified helpers are shared by the relevant detachment code.

## Fold-in decision

Promote to correctness loop and commit as a scoped production fix with tests. This strengthens the auditable mutation boundary by ensuring imported formula-binding metadata is either detached or journaled instead of silently ignored due to sheet-name casing.

## Next question

Can the safe-open timing probe run from a tracked-clean tree now, or are more in-flight correctness changes still landing?
