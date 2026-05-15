# Verify Pivot Sheet Ref Casing

## Question

Does the verifier falsely flag pivot tables and pivot cache sources when sheet and table metadata casing differs from the workbook model?

## Hypothesis

It should not. Pivot metadata imported from XLSX parts can preserve different casing from the workbook model, but sheet/table identity is still the same for Ascend's structural checks.

## External sources checked

- Microsoft worksheet references define sheet/cell references as worksheet identity plus cell block, not a raw string comparison contract: https://learn.microsoft.com/en-us/office/client-developer/excel/worksheet-references
- Microsoft structured references document table names and column names as symbolic formula references: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Open Packaging Conventions model workbook evidence as related package parts, which is why verifier diagnostics must distinguish stale metadata from harmless lexical casing drift: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/packages-overview

## Why this matters to Ascend

The package-action and safe-open claims both depend on verifier issues being meaningful. A false `pivot-table-sheet-missing` or `pivot-cache-source-table-missing` issue caused only by casing drift would make audit evidence noisy and less trustworthy.

## Probe/implementation

Commit `de8f4128 fix(verify): classify pivot sheet refs case-insensitively` updated `packages/verify/src/checker.ts` so pivot integrity checks compare:

- pivot table owner sheets case-insensitively;
- pivot cache source sheets case-insensitively;
- pivot cache source table names case-insensitively;
- source table sheet mismatch checks case-insensitively.

It added a verifier regression with:

- workbook sheet `Sheet1`;
- table `SalesTable`;
- pivot table sheet `sheet1`;
- pivot cache source sheet `sheet1`;
- pivot cache source name `salestable`.

## Results

The regression expects no `pivot-integrity` or `pivot-source-integrity` issues from casing drift alone. The latest changed-test run after this commit passed:

```bash
bun run test:changed
```

Result: 4379 pass, 1 skip, 0 fail across 168 files.

## Confidence

High for pivot table owner sheet, pivot cache source sheet, and source table name checks. Medium for every represented metadata class because this was targeted to pivot/slicer verification.

## Fold-in decision

Folded into the correctness loop as a verifier false-positive fix. It supports auditable package-part mutation by making pivot audit failures more meaningful.

## Next question

Do engine sheet operations remove, rename, and clone chart/pivot metadata case-insensitively when sheet ownership casing drifts?
