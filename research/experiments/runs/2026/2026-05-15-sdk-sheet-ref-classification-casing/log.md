# SDK sheet ref classification casing

## Question

Do SDK journal and engine shift helpers classify sheet-qualified formula refs case-insensitively when deciding whether a mutation touches a sheet, deleted axis, or moved range?

## Hypothesis

Yes. After the formula rewrite casing fixes, the remaining journal/ref-classification helpers should also use case-insensitive sheet comparison so `sheet1!A1` is treated as a reference to `Sheet1`.

## External sources checked

- Microsoft Open XML formulas overview describes formulas and cross-sheet references in workbook XML: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- Microsoft names-in-formulas documentation describes workbook/worksheet-scoped symbols and formula references: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft worksheet rename documentation frames sheet names as user-visible workbook identifiers: https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9

## Why this matters to Ascend

Mutation journals and structural rewrite previews need to classify affected formulas consistently. A case-sensitive comparison can make a journal look exact or unrelated when the workbook formula actually points at the mutated sheet.

## Probe/implementation

Committed production change:

```text
576b77a1 fix(sdk): classify sheet refs case-insensitively
```

Touched paths:

- `packages/engine/src/structural/formula-rewrite.ts`
- `packages/sdk/src/journal.ts`
- `packages/engine/src/operations.test.ts`
- `packages/sdk/src/interactive-contract.test.ts`

The change makes:

- row/column shift target checks compare sheet names case-insensitively;
- journal sheet-reference classification compare owner and explicit sheets case-insensitively;
- journal deleted-axis and moved-range overlap classification compare sheets case-insensitively.

Validation:

```bash
bun test packages/engine/src/operations.test.ts -t "case-insensitive|insertRows|deleteRows|shift|moveRange"
bun test packages/sdk/src/interactive-contract.test.ts packages/sdk/src/journal-exactness.test.ts -t "case-insensitive|delete|move|sheet"
bunx biome check packages/engine/src/structural/formula-rewrite.ts packages/sdk/src/journal.ts
bunx tsc --build
bun run test:changed
```

## Results

- Engine focused structural/casing suite passed: 56 tests, 0 fail.
- SDK journal focused suite passed: 59 tests, 0 fail.
- Biome passed.
- TypeScript build passed.
- `bun run test:changed` passed: 4378 pass, 1 skip, 0 fail across 168 files.
- No formula-assist rename surface was added; this is operation/journal-owned classification.

## Confidence

High for the changed helpers and journal/operation paths exercised by the focused tests. Medium for every metadata surface because the proof is helper-backed rather than a bespoke case test for each carrier.

## Fold-in decision

Promote to correctness loop and keep as production fix. This strengthens auditable mutation and formula-reference rewrite claims without promoting formula language-service rename.

## Next question

Can the practical latency contract phase-profile change finish cleanly so the safe-open timing probe can finally run from an empty tracked baseline?
