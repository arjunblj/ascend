# Verify shared master casing

## Question

Does Verify accept shared-formula member metadata when master references are sheet-qualified with different casing from the workbook sheet name?

## Hypothesis

Yes. The verifier already normalizes sheet-qualified binding ranges; it should also compare parsed shared-formula master cells case-insensitively so `Data!E1` and `data!$E$1` identify the same master cell.

## External sources checked

- Microsoft Open XML formulas overview describes formulas, references, and worksheet XML storage: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- Open XML SDK `CellFormula` docs describe shared formula type, `ref`, and shared-index attributes: https://learn.microsoft.com/es-es/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1
- Apache POI formula type docs describe shared formulas as optimized formula metadata: https://poi.apache.org/apidocs/3.17/org/apache/poi/ss/formula/FormulaType.html

## Why this matters to Ascend

Formula-binding integrity diagnostics are part of Ascend's correctness proof. They must reject stale imported metadata without creating false positives for casing drift in producer-authored sheet-qualified refs.

## Probe/implementation

Updated `packages/verify/src/checker.ts` so shared-formula master equality uses `sameSheetName` after parsing both master refs. Extended the existing case-insensitive formula-binding test in `packages/verify/src/verify.test.ts` with:

- master formula metadata at `Data!E1`;
- member formula metadata with `masterRef: "data!$E$1"`;
- expectation that `formula-binding-integrity` emits no issue.

Committed production change:

```text
d942a94a fix(verify): match shared masters case-insensitively
```

Validation:

```bash
bun test packages/verify/src/verify.test.ts -t "case-insensitive sheet-qualified formula binding refs"
bun test packages/verify/src/verify.test.ts
bunx biome check packages/verify/src/checker.ts packages/verify/src/verify.test.ts
bunx tsc --build
bun run test:changed
```

## Results

- Targeted verifier test passed.
- Full verify test file passed: 147 tests, 0 fail.
- Biome passed for both touched files.
- TypeScript build passed.
- `bun run test:changed` passed: 3988 pass, 1 skip, 0 fail across 136 files.

## Confidence

High for shared-formula master/member equality in Verify. This is a single comparison fix with direct regression coverage.

## Fold-in decision

Promote to correctness loop and commit. This completes the verifier side of the case-insensitive formula-binding metadata fix family.

## Next question

Can the safe-open timing probe be rerun after this verify fix is committed from a tracked-clean tree?
