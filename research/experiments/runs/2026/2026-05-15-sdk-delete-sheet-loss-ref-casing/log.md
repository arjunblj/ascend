# SDK Delete Sheet Loss Ref Casing

## Question

Do SDK delete-sheet journals report dependent metadata refs when represented sheet, chart, and pivot metadata use different casing from the deleted sheet?

## Hypothesis

They should. Lossy delete-sheet evidence must identify dependent package metadata by sheet identity, not exact string casing.

## External sources checked

- Microsoft worksheet references define worksheet/cell references as symbolic workbook targets: https://learn.microsoft.com/en-us/office/client-developer/excel/worksheet-references
- Microsoft structured references document workbook/table references as symbolic formula metadata: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Open Packaging Conventions frame workbook evidence as related package parts, which makes dependent chart/pivot accounting part of write proof: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/packages-overview

## Why this matters to Ascend

The auditable package-part mutation claim depends on journals surfacing lossy evidence before commit. If `deleteSheet('Data')` misses `data` chart/pivot ownership, the proof undercounts dependent package metadata.

## Probe/implementation

Commit `fbddc12d fix(sdk): journal sheet metadata refs case-insensitively` updated `packages/sdk/src/journal.ts` so delete-sheet loss refs compare:

- sheet-scoped defined-name owners;
- sibling sheet formula owners;
- chart owners;
- pivot table owners;
- pivot cache source sheets.

The regression in `packages/sdk/src/interactive-contract.test.ts` seeds chart, pivot table, and pivot cache metadata with lowercase `data` while deleting `Data`.

## Results

Focused validation passed:

```bash
bun test packages/sdk/src/interactive-contract.test.ts -t "delete sheet journals surface lost sheet contents"
```

The full changed-test gate after the casing cluster passed:

```bash
bun run test:changed
```

Latest result during this checkpoint: 4379 pass, 1 skip, 0 fail across 168 files.

## Confidence

High for delete-sheet journal refs covered by the regression. Medium for every package metadata kind because this remains a targeted casing audit.

## Fold-in decision

Folded into correctness as journal evidence hygiene for auditable package-part mutation.

## Next question

Do chart, pivot, and connection operations use the same sheet identity semantics when selecting metadata by sheet?
