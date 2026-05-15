# Engine Connection Selector Casing

## Question

Can `setConnectionRefresh` select query-table connection metadata by sheet case-insensitively?

## Hypothesis

Yes. Query-table connection parts are workbook metadata owned by a worksheet, and an operation that names `Sheet1` should match stored owner `sheet1`.

## External sources checked

- Microsoft query tables document worksheet-owned external data ranges: https://learn.microsoft.com/en-us/office/vba/api/excel.querytable
- Microsoft worksheet references define sheet/cell targets as worksheet references: https://learn.microsoft.com/en-us/office/client-developer/excel/worksheet-references
- Open Packaging Conventions describe connection/query-table metadata as package parts and relationships: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/packages-overview

## Why this matters to Ascend

Connection refresh metadata is a trust boundary for external data. Selector misses caused by sheet-name casing drift could prevent an agent from marking stale external data for refresh or surfacing warnings before write.

## Probe/implementation

Commit `691576e2 fix(engine): select connection sheet owners case-insensitively` updated `packages/engine/src/operations/connection-ops.ts` so sheet selector matching uses a case-insensitive optional sheet comparison.

The regression uses connection/query-table sheet metadata with lowercase casing and selects it through the visible sheet name.

## Results

The final changed-test checkpoint after this commit passed:

```bash
bun run test:changed
```

Result: 4379 pass, 1 skip, 0 fail across 168 files.

## Confidence

Medium-high for connection refresh selector behavior. Medium for all query-backed table workflows because refresh correctness still depends on Excel or another connection-aware engine.

## Fold-in decision

Folded into correctness as a tiny selector fix. Do not promote a stronger query refresh claim.

## Next question

With in-flight casing fixes committed, can research return to claim stewardship and avoid adding new production surfaces unless a top-two proof gate requires it?
