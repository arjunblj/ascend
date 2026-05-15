# Formula sheet rewrite casing

## Question

Do formula sheet-reference rewrite helpers update sheet-qualified references when formula text casing differs from the workbook sheet name?

## Hypothesis

Yes. Excel-visible sheet identity should be treated case-insensitively by rename/copy rewrite helpers. A formula like `sheet1!A1+1` should rewrite to `Data!A1+1` when `Sheet1` is renamed to `Data`.

## External sources checked

- Microsoft Open XML formulas overview describes formulas and cross-sheet references in SpreadsheetML: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- Microsoft worksheet rename documentation lists sheet-name constraints and frames sheet names as user-visible workbook identifiers: https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
- Microsoft names-in-formulas documentation describes workbook and worksheet scope for symbolic formula references: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1

## Why this matters to Ascend

Safe mutation planning depends on rewriting the formulas that Excel would consider references to the edited sheet. If the rewrite path is case-sensitive, imported formulas with producer-specific casing can retain stale sheet names after rename or copy operations.

## Probe/implementation

Updated `packages/engine/src/structural/formula-rewrite.ts` so sheet-reference detection and rewrite compare sheet names through a local case-insensitive helper. Covered:

- explicit cell/range refs retargeted during copy-sheet range rewrites;
- `formulaAstReferencesSheet`;
- `rewriteSheetName` for cell refs, range refs, whole-row/whole-column refs, and 3D sheet spans.

Committed production change:

```text
a6ca6b94 fix(engine): rewrite sheet refs case-insensitively
```

Updated `packages/engine/src/operations.test.ts` so:

- `copySheet` retargets `sheet1!A1+1`;
- `renameSheet` rewrites both a defined name and formula text that use `sheet1`.

Validation:

```bash
bun test packages/engine/src/operations.test.ts -t "renameSheet updates sheet name|copySheet retargets copied sheet-qualified"
bunx biome check packages/engine/src/operations.test.ts packages/engine/src/structural/formula-rewrite.ts
bunx tsc --build
bun run test:changed
```

## Results

- Engine targeted tests passed: 2 tests, 15 assertions.
- Biome passed after formatting.
- TypeScript build passed.
- `bun run test:changed` passed: 4378 pass, 1 skip, 0 fail across 168 files.
- No formula language-service rename surface was added; this is operation-owned sheet rewrite behavior.

## Confidence

High for rename-sheet and copy-sheet explicit formula references. Medium for every formula metadata rewrite path because this probe targets the shared AST helper but does not enumerate every metadata carrier separately.

## Fold-in decision

Promote to correctness loop and commit with tests. This is a legitimate production fold-in from the case-insensitive formula-reference evidence family.

## Next question

Should formula metadata rewrite tests add x14 validation/conditional-format sheet-ref casing examples, or is the shared AST helper enough for this loop?
