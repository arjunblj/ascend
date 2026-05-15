# Engine Chart Pivot Sheet Ownership Casing

## Question

Do engine sheet operations remove, rename, and clone chart/pivot metadata when the metadata's owning sheet casing differs from the operation sheet casing?

## Hypothesis

They should. Delete, rename, and copy operations should use sheet identity semantics for chart and pivot ownership, or preserved package metadata can be orphaned, skipped, or omitted from copied sheets.

## External sources checked

- Microsoft worksheet references describe worksheet-scoped references as cell blocks owned by worksheets: https://learn.microsoft.com/en-us/office/client-developer/excel/worksheet-references
- Microsoft structured references document symbolic table references whose casing may differ from the displayed table or sheet metadata: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Open Packaging Conventions describe workbook objects as package parts and relationships, which makes stale chart/pivot metadata a package-accounting problem after sheet operations: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/packages-overview

## Why this matters to Ascend

Ascend's preservation-first mutation story includes visual and analytical metadata. If `deleteSheet('Sheet2')` leaves a pivot or chart whose owner is `sheet2`, or `copySheet('Sheet1')` skips a chart whose owner is `sheet1`, audit and reopened workbook state can diverge from user intent.

## Probe/implementation

Commit `71e65deb fix(engine): match sheet metadata owners case-insensitively` updated `packages/engine/src/operations/sheet-ops.ts` so these paths use case-insensitive sheet ownership checks:

- removed pivot names during `deleteSheet`;
- pivot-table metadata removal for a deleted sheet;
- chart-part removal for a deleted sheet;
- chart owner rewrite during `renameSheet`;
- chart cloning during `copySheet`.

Existing operation regressions already seed lower-case chart/pivot owner metadata for:

- `deleteSheet` pivot metadata removal;
- `renameSheet` chart owner and source refs;
- `deleteSheet` chart-part removal;
- `copySheet` chart cloning and source ref retargeting.

## Results

Focused validation passed:

```bash
bun test packages/engine/src/operations.test.ts -t "deleteSheet removes sheet-scoped names|chart ownership|deleteSheet removes chart|copySheet duplicates visual"
bunx biome check packages/engine/src/operations/sheet-ops.ts packages/engine/src/operations.test.ts
bunx tsc --build
bun run test:changed
```

Results:

- focused engine tests: 4 pass, 0 fail;
- Biome: checked 2 files, no fixes applied;
- typecheck: passed;
- changed tests: 4379 pass, 1 skip, 0 fail across 168 files.

## Confidence

High for the covered delete, rename, and copy sheet metadata ownership paths. Medium for every represented metadata class because this is a targeted ownership fix, not a full casing audit.

## Fold-in decision

Folded in as a tiny correctness fix. Do not promote a new surface or claim; this only tightens the auditable mutation proof boundary.

## Next question

After committing this tiny fix, refresh the claim board and stop implementation promotion unless a top-two proof owner explicitly needs another scoped correction.
