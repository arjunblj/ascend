# Engine Chart Selector Casing

## Question

Can `setChartSeriesSource` select chart metadata by sheet case-insensitively while preserving stable inverse journal identity?

## Hypothesis

Yes. A chart owned by `sheet1` should be selectable through an operation that names `Sheet1`, but inverse journals should still prefer stable `partPath` identity when available.

## External sources checked

- Microsoft worksheet references describe sheet/cell references as symbolic workbook references: https://learn.microsoft.com/en-us/office/client-developer/excel/worksheet-references
- Open Packaging Conventions describe chart parts as package parts connected by relationships: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/packages-overview
- Microsoft Open XML chart documentation exposes chart parts as workbook package objects rather than sheet-name strings alone: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-charts

## Why this matters to Ascend

Chart source edits are preservation-sensitive. Agents should be able to select chart metadata using the visible sheet name even when parsed package metadata preserves different casing, and rollback evidence must still be exact.

## Probe/implementation

Commit `a267762b fix(engine): select chart sheet owners case-insensitively` updated:

- `packages/engine/src/operations/visual-ops.ts` sheet selector filtering for chart parts;
- `packages/sdk/src/journal.ts` chart series preimage filtering;
- engine and SDK regressions with chart owner `sheet1` and operation sheet `Sheet1`.

## Results

Focused validation passed:

```bash
bun test packages/engine/src/operations.test.ts -t "setChartSeriesSource updates parsed chart source refs"
bun test packages/sdk/src/interactive-contract.test.ts -t "journal inverse ops restore chart series source refs"
bunx biome check packages/engine/src/operations/visual-ops.ts packages/engine/src/operations.test.ts packages/sdk/src/journal.ts packages/sdk/src/interactive-contract.test.ts
```

Changed tests later passed with 5089 pass, 1 skip, 0 fail when `core` was also in the affected graph, and 4379 pass, 1 skip, 0 fail after the later benchmark-only checkpoint.

## Confidence

High for chart selection by sheet plus inverse journal exactness. Medium for every visual metadata selector until a full visual-ops selector audit is complete.

## Fold-in decision

Folded into correctness. This is not a new chart API; it makes existing chart selection and journal rollback evidence less brittle.

## Next question

Do pivot item selectors and connection refresh selectors use the same case-insensitive sheet identity semantics?
