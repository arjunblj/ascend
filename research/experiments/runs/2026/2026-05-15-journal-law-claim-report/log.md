# Journal Law Claim Report

Date: 2026-05-15

## Question

Can the deterministic journal-law harness emit release-facing claim wording without promoting style/table-style exactness or full property-based testing?

## Hypothesis

Yes. A compact claim report can separate allowed wording, exact-law families, lossy issue reasons, and "do not promote yet" boundaries from the raw case table. That gives correctness/product loops a proof artifact without adding a public SDK/CLI/API/MCP surface.

## External sources checked

- fast-check model-based testing documents command sequences, `check`/`run`, shrinkers, and `replayPath`, which defines what Ascend still lacks before claiming shrinkable property-based journal laws: https://fast-check.dev/docs/advanced/model-based-testing/
- Microsoft Open XML SDK `CellStyleFormats` docs explain that style records reference shared format collections, supporting the decision to keep style restoration as package-state sensitive: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-3.0.1
- OOXML Info's Styles Part page identifies `styleSheet` as the styles package part, reinforcing the boundary that cell style edits are not only cell-local evidence: https://ooxml.info/docs/12/12.3/12.3.20/

## Why this matters to Ascend

Research has already produced the raw proof harness. The missing product-shaped artifact is claim stewardship: what Ascend can safely say, what the proof covers, and what must stay out of release wording. Without that split, a useful correctness harness can accidentally become an overbroad product claim.

## Probe/implementation

Folded a claim-report mode into `fixtures/benchmarks/journal-law-proof.ts`:

- `journalLawClaimReport(result)` returns structured allowed claim wording, proof status, exact families, lossy issue reasons, "do not promote yet" boundaries, and next proof.
- `journalLawClaimReportMarkdown(report)` renders the report.
- CLI support: `bun run fixtures/benchmarks/journal-law-proof.ts --claim-report` and `--claim-report --json`.
- The case result now records operation family names so exact-family coverage can be reported without guessing from lossy cases.
- Added test coverage in `fixtures/benchmarks/journal-law-proof.test.ts`.

## Results

Probe commands:

```bash
bun run fixtures/benchmarks/journal-law-proof.ts --claim-report
bun run fixtures/benchmarks/journal-law-proof.ts --claim-report --json
bun test fixtures/benchmarks/journal-law-proof.test.ts
```

Observed:

- Proof status: passed.
- Exact law cases: 56.
- Lossy boundaries: 7.
- Exact operation families include cells, formulas, comments, hyperlinks, freeze panes, data validations, conditional formats, row/column/page metadata, workbook/document properties, workbook views, calc settings, workbook protection, and theme.
- Lossy issue reasons remain explicit for package-part preservation, table metadata, data-validation metadata order/default/duplicate issues, and conditional-format order/duplicate issues.

## Confidence

High that this improves claim discipline for the current deterministic artifact. Medium for the broader journal-law claim because the harness is still not shrinkable and style/table-style exactness remains intentionally unpromoted.

## Fold-in decision

Fold into the correctness proof harness as a report mode. Do not add it to the top release proof index yet.

## Next question

Should the next correctness loop add fast-check model-based generation with replayable shrinking, or keep journal laws deterministic and move to another higher-ranked portfolio claim?
