# SDK Metadata Ref Casing

## Question

Do SDK mutation journals classify chart and pivot metadata refs case-insensitively when the represented package metadata uses a different sheet-name casing than the workbook model?

## Hypothesis

Yes. Sheet-qualified references in represented metadata should be matched by sheet identity, not string casing, or structural delete journals can miss lossy package metadata.

## External sources checked

- Microsoft worksheet references describe worksheet references as references to sheet-local cell blocks, with separate internal/external reference identity: https://learn.microsoft.com/en-us/office/client-developer/excel/worksheet-references
- Microsoft structured references show Excel table and column references as symbolic workbook references rather than plain strings: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Open Packaging Conventions model packages as parts and relationship graphs, reinforcing why represented chart/pivot metadata must stay explainable after mutation: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/packages-overview

## Why this matters to Ascend

The auditable package-part mutation claim depends on journal issue reasons that do not silently miss represented package metadata. If a chart series or pivot cache points at `sheet1!A2:A2` while the in-memory sheet is `Sheet1`, deleting that row should still produce the same lossy inverse evidence.

## Probe/implementation

Commit `4f2fabcd fix(sdk): classify metadata refs case-insensitively` updated `packages/sdk/src/journal.ts` so `refTextOverlapsAffected()` compares sheet names with the existing case-insensitive helper.

The regression changed `packages/sdk/src/interactive-contract.test.ts` to seed:

- chart series refs as `sheet1!$A$2:$A$2`;
- pivot cache source sheets as `sheet1`.

## Results

The structural delete journal still reports package metadata refs as lossy, including:

- `chart:xl/charts/chart1.xml:series:0:valueRef`
- `pivotCache:xl/pivotCache/pivotCacheDefinition1.xml:sourceRef`
- `pivotTable:xl/pivotTables/pivotTable1.xml:locationRef`

This closes a small correctness gap inside the package-action proof story. It does not promote a new journal surface.

## Confidence

High for chart and pivot metadata overlap checks. Medium for all workbook metadata classes because this was a targeted fix after a cluster of sheet-casing bugs, not a full workbook-wide audit.

## Fold-in decision

Folded into the correctness loop as a tiny journal fix supporting auditable package-part mutation.

## Next question

Do engine-owned hyperlink locations have the same sheet-casing issue during row/column shifts and sheet rename?
