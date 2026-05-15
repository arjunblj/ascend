# Journal Validation Boundary Sweep

## Question

Which remaining journal exactness paths were still claiming exact rollback for public operations that the engine rejects or cannot address?

## Hypothesis

The highest-value correctness work is not adding new mutation surfaces; it is making existing journal proof refuse invalid operation shapes consistently. Workbook metadata, formula syntax, pivot cache selectors, defined-name deletion, print-area ranges, sheet-name creation, and table-name creation are likely boundary cases because their public operation shapes include optional selectors, editable fields, ranges, or Excel naming constraints.

## External sources checked

- Microsoft Open XML workbook properties class family: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookproperties?view=openxml-3.0.1
- Microsoft Excel formula overview: https://support.microsoft.com/en-us/office/overview-of-formulas-34519a4e-1e8d-4f4b-84d4-d642c4f63263
- Microsoft Open XML pivot cache `WorksheetSource`: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.worksheetsource?view=openxml-3.0.1
- Microsoft Excel defined names overview: https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64
- Microsoft print area guidance: https://support.microsoft.com/en-us/office/set-or-clear-a-print-area-on-a-worksheet-27048af8-a321-416d-ba1b-e99ae2182a7e
- Microsoft worksheet naming limits: https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
- Microsoft Excel table naming guidance: https://support.microsoft.com/en-us/office/rename-an-excel-table-017b31f8-3b46-4415-8f2c-7b7c27d5081b

## Why this matters to Ascend

Ascend's auditable mutation claim is only credible if journal exactness is rejection-first. A journal that says "exact" for an edit the engine rejects is worse than no journal because it tells agents and reviewers the rollback story is safer than it is.

## Probe/implementation

Local invalid-operation probes compared engine apply results with `analyzeMutationJournalExactness`. The sweep folded in four correctness fixes:

- Workbook metadata operations now classify rejected value shapes as `UNSUPPORTED_VALUE`.
- Formula journal validation now covers invalid `setFormula`/`fillFormula` text as unsupported formula values.
- Pivot cache edits now require a public selector, at least one editable field, and valid source refs before inverse emission.
- Missing `deleteDefinedName` targets now classify as unsupported values instead of exact no-op rollback.
- Invalid `setPrintArea` ranges now fail in the engine before writing defined-name metadata and classify as unsupported defined-name values in the journal.
- Invalid `addSheet`, `copySheet`, and `renameSheet` target names now classify as sheet topology issues before inverse emission.
- Invalid `createTable` and `renameTable` target names now classify as unsupported table values before inverse emission.

## Results

Validation:

```bash
bun test packages/sdk/src/journal-exactness.test.ts
bunx biome check packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts fixtures/benchmarks/agent-workflow.test.ts
bunx tsc --build
```

All passed. `journal-exactness.test.ts` now has explicit boundary cases for rejected workbook metadata, invalid formulas, rejected pivot cache edits, missing defined-name deletes, invalid print-area ranges, invalid target sheet names, and invalid target table names.

## Confidence

High for the covered operation shapes. Medium for the full public operation set because the sweep was targeted at newly discovered validation gaps, not a generated exhaustive invalid-operation corpus.

## Fold-in decision

Promote to correctness loop as proof hygiene. Do not market broader mutation compatibility from this sweep; it only prevents false exactness claims for rejected operation shapes.

## Next question

Can invalid-operation probes be promoted into a generated conformance harness so future public operations cannot accidentally add an exact journal for a rejected edit?
