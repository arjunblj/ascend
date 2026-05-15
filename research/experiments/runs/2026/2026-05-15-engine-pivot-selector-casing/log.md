# Engine Pivot Selector Casing

## Question

Can `setPivotFieldItem` select a pivot table by sheet case-insensitively when the stored pivot owner casing differs from the operation sheet?

## Hypothesis

Yes. A pivot table whose metadata owner is `sheet1` should match `sheet: "Sheet1"` and return modified sheet evidence using the stored owner identity.

## External sources checked

- Microsoft pivot tables are workbook metadata bound to worksheet output locations: https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576
- Microsoft worksheet references describe sheet/cell targets as worksheet-owned references: https://learn.microsoft.com/en-us/office/client-developer/excel/worksheet-references
- Open Packaging Conventions model pivot tables and caches as package parts whose relationships must remain explainable: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/packages-overview

## Why this matters to Ascend

Pivot operations are high-risk because saved output can be stale until Excel refreshes cache state. Selector misses caused by casing drift would make agent workflows brittle and could hide refresh warnings.

## Probe/implementation

Commit `944d2d7f fix(engine): select pivot sheet owners case-insensitively` updated `packages/engine/src/operations/pivot-ops.ts` so the pivot-table sheet selector compares sheet names case-insensitively.

Regression coverage seeds pivot owner `sheet1` and applies:

```ts
{ op: 'setPivotFieldItem', sheet: 'Sheet1', fieldIndex: 0, itemIndex: 0, hidden: null }
```

## Results

Focused validation passed:

```bash
bun test packages/engine/src/operations.test.ts -t "replaceImage swaps media bytes|setPivotFieldItem"
bunx biome check packages/engine/src/operations/pivot-ops.ts packages/engine/src/operations.test.ts
bunx tsc --build
bun run test:changed
```

Result: 4379 pass, 1 skip, 0 fail across 168 files.

## Confidence

High for `setPivotFieldItem` sheet selector behavior. Medium for broader pivot workflows because cache selector and refresh semantics have separate proof surfaces.

## Fold-in decision

Folded into correctness as a tiny selector fix for existing pivot operations.

## Next question

Do connection refresh selectors match query-table sheet owners case-insensitively?
