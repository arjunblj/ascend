# Pivot Cache Journal Boundary

## Question

Can `setPivotCache` journals claim exact rollback when the edit is missing a selector, missing an editable field, or carries an invalid source range?

## Hypothesis

No. Pivot cache source edits need an addressable cache and a valid source reference. If those public operation fields are absent or invalid, the journal should classify the attempted edit as an unsupported value instead of emitting a lossy or misleading inverse.

## External sources checked

- Microsoft Open XML `WorksheetSource` pivot cache source documentation: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.worksheetsource?view=openxml-3.0.1
- Microsoft Open XML `CacheSource` pivot cache source description: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cachesource?view=openxml-3.0.1
- Microsoft Excel PivotTable source data guidance: https://support.microsoft.com/en-au/office/change-the-source-data-for-a-pivottable-afd93524-f7de-432c-84d0-3896fbbc2577

## Why this matters to Ascend

Pivot cache metadata is a high-risk workbook feature for preservation-first XLSX. Auditable mutation claims need journal evidence that distinguishes "we can restore this cache source field" from "the requested public edit was not even valid enough to build an exact inverse."

## Probe/implementation

- Probed invalid `setPivotCache` variants against engine apply behavior and journal exactness.
- Folded in value validation before inverse emission:
  - missing selector
  - selector with no editable fields
  - invalid `sourceRef`
- Added `value-unsupported` to the `pivot-caches` exactness matrix.
- Added regression coverage for all three rejected forms.

## Results

Validation:

```bash
bun test packages/sdk/src/journal-exactness.test.ts
bunx biome check packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts fixtures/benchmarks/agent-workflow.test.ts
bunx tsc --build
```

All passed. The focused pivot-cache test reports `exact=false`, `surface=pivot-caches`, `reason=value-unsupported`, and no matrix violation for each rejected operation shape.

## Confidence

High for the three public validation boundaries covered by the regression. Medium for deeper imported pivot cache semantics because this pass did not attempt to understand every cache definition or records payload.

## Fold-in decision

Promote to correctness loop as journal proof hygiene. This does not promote new pivot editing claims; it only prevents invalid pivot cache edits from being represented as exact rollback evidence.

## Next question

Can the remaining pivot and slicer journal operations expose a shared selector/value validation helper without hiding feature-specific boundaries?
