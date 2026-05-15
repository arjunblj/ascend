# Journal Law Package State Style Boundary

Date: 2026-05-15

## Question

Can the journal-law proof cover package-state replacements while refusing to call style and table-style package gaps exact?

## Hypothesis

Yes. Workbook/document properties, workbook views, calc settings, workbook protection, and theme metadata have representable replacement inverses. Cell style and table style edits still touch package/table metadata that public inverse operations cannot fully restore, so they should remain lossy boundaries.

## External sources checked

- Microsoft Open XML SDK `Workbook` docs list workbook-level children including calculation properties and workbook views: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbook?view=openxml-3.0.1
- Microsoft Open XML SDK `CellStyleFormats` docs describe style format records as references into shared style collections, which makes style restoration a package-state problem rather than only a cell problem: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-3.0.1
- Microsoft Open XML SDK `Stylesheet` docs identify the stylesheet as the package part holding fonts, fills, borders, cell formats, differential formats, and cell styles: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.stylesheet?view=openxml-3.0.1
- fast-check model-based testing remains the next candidate when this deterministic proof needs shrinkable operation sequences: https://fast-check.dev/docs/advanced/model-based-testing/

## Why this matters to Ascend

The product-shaped claim is auditable mutation, not "all undo works." Package-level state is where spreadsheet tools often lose trust: document properties, themes, workbook views, styles, and table style metadata are outside ordinary cell semantics. The proof must show which package-state edits are exact and which still have honest preservation gaps.

## Probe/implementation

Expanded `fixtures/benchmarks/journal-law-proof.ts` with:

- exact package-state cases for workbook/document properties, workbook view plus calc settings plus workbook protection, and theme replacement;
- lossy style/table boundaries for style-number-format package preservation and table-style metadata replacement;
- stronger test assertions in `fixtures/benchmarks/journal-law-proof.test.ts` for package-state operation families and lossy issue reasons.

## Results

Probe:

```bash
bun run fixtures/benchmarks/journal-law-proof.ts --json
```

Observed:

- 63 total cases.
- 56 exact law cases.
- 7 lossy boundary cases.
- 0 failures.
- New exact package-state families: `setDocumentProperties`, `setWorkbookProperties`, `setWorkbookView`, `setCalcSettings`, `setWorkbookProtection`, and `setTheme`.
- New lossy boundaries: `package-parts:package-part-preservation=2` for style/number-format edits and `tables:table-metadata=1` for table style edits.

## Confidence

Medium-high for these package-state replacement boundaries. Medium overall because style/table-style gaps are still lossy and the harness remains deterministic rather than shrinkable.

## Fold-in decision

Fold into the correctness proof harness. Do not promote style/table-style exact claims; keep them in the "do not promote yet" list until public inverse operations can restore the relevant package/table metadata exactly.

## Next question

Should the next proof step be shrinkable model-based journal laws, or a release-facing journal-law report that turns the exact/lossy table into claim-safe wording?
