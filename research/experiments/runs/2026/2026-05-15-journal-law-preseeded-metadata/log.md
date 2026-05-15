# Journal Law Preseeded Metadata

Date: 2026-05-15

## Question

Can pre-seeded worksheet metadata broaden exact inverse journal-law coverage without hiding known creation-loss or selector-loss boundaries?

## Hypothesis

Yes. Row/column layout, sheet protection, tab color, page setup, and print area should be exact when the metadata already exists and the operation is a replacement. Creation and ambiguous selector cases should remain separate lossy boundaries.

## External sources checked

- Microsoft Open XML SDK `Row` docs identify worksheet row metadata as an `x:row` element with cell children and row-level attributes: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row?view=openxml-3.0.1
- Microsoft Open XML SDK `Column` docs identify column width/format metadata as an `x:col` leaf element: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-3.0.1
- Microsoft Open XML SDK `PageSetup` docs identify page setup as the `x:pageSetup` worksheet element: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pagesetup?view=openxml-3.0.1
- Microsoft Open XML SDK `SheetProtection` docs identify sheet protection as the `x:sheetProtection` worksheet element: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetprotection?view=openxml-3.0.1
- fast-check model-based testing remains the likely next shape for shrinkable operation sequences: https://fast-check.dev/docs/advanced/model-based-testing/

## Why this matters to Ascend

Auditable mutation planning needs to distinguish exact replacement laws from operations that lose package defaults, ordering, or selector identity. If the proof harness only tests easy cell-like operations, the journal-law claim is too narrow. If it treats creation cases as exact, the claim becomes dishonest. Pre-seeding metadata tests the useful middle.

## Probe/implementation

Expanded the tracked proof harness in `fixtures/benchmarks/journal-law-proof.ts` with exact-law cases that seed metadata first, then replace it under journal tracking:

- existing row height plus hidden row replacement;
- existing column width plus hidden column replacement;
- existing sheet protection replacement;
- existing tab color replacement;
- existing page setup plus print-area replacement.

Updated `fixtures/benchmarks/journal-law-proof.test.ts` to require the broader exact count and operation-family coverage.

## Results

Validation:

```bash
bun run fixtures/benchmarks/journal-law-proof.ts
bun test fixtures/benchmarks/journal-law-proof.test.ts
bunx biome check --write fixtures/benchmarks/journal-law-proof.ts fixtures/benchmarks/journal-law-proof.test.ts
```

Observed:

- 58 total cases.
- 53 exact law cases.
- 5 lossy boundary cases.
- 0 failures.
- New exact families covered: `setRowHeight`, `hideRows`, `setColWidth`, `hideCols`, `setSheetProtection`, `setTabColor`, `setPageSetup`, and `setPrintArea`.
- Lossy boundaries stayed explicit for metadata order and duplicate selector cases.

## Confidence

Medium. The replacement boundary is now clearer and tested, but the harness is still deterministic and not shrinkable. Package-state and style operations remain outside the exact proof.

## Fold-in decision

Fold into the correctness proof harness. Do not promote property-style journal laws to the release proof index yet.

## Next question

Should the next correctness proof add shrinkable `fast-check` model-based generation, or should it first add package-state/style exact cases with the same pre-seeded replacement discipline?
