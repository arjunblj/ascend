# Safe open proof bundle boundary

## Question

What exactly should Ascend prove next for the product-shaped claim "safe unknown workbook opening," and does the current implementation need another production surface?

## Hypothesis

The implementation does not need another surface. SDK, CLI, API, and MCP already expose `open-plan`; the missing proof is a fixture-backed report that combines recommendation, risk reasons, latency contrast, and honest malformed-package handling.

## External sources checked

- [openpyxl usage docs](https://openpyxl.readthedocs.io/en/3.0/usage.html) document `keep_vba` and warn that unsupported items such as images and charts can be lost when opening and saving existing files.
- [SheetJS write options](https://docs.sheetjs.com/docs/api/write-options/) frame writer behavior around data-preservation options, which is a contrast point for Ascend's pre-hydration routing claim.
- [ExcelJS project docs](https://exceljs.org/) emphasize unified read/write and streaming support, but do not frame unknown workbook opening as a package-risk routing step.
- [LibreOffice basics](https://books.libreoffice.org/en/GS74/GS7401-LOBasics.html) document Safe Mode as an application recovery mode, a useful contrast because Ascend's "safe opening" is package-level routing before workbook hydration, not office-suite recovery or malware scanning.

## Why this matters to Ascend

The claim ladder ranked safe unknown workbook opening first because it is already credible and product-shaped. The next proof must prevent overclaiming: pre-hydration routing is valuable, but it is not sandboxing, malware detection, or a guarantee that malformed packages can be inspected.

## Probe/implementation

- Inspected current SDK, CLI, API, and MCP open-plan implementations.
- Ran a local Bun probe over public fixtures:
  - `fixtures/xlsx/poi/SampleSS.xlsx`
  - `fixtures/xlsx/calamine/vba.xlsm`
  - `fixtures/xlsx/poi/ExcelPivotTableSample.xlsx`
  - `fixtures/xlsx/libreoffice/activex_checkbox.xlsx`
  - `fixtures/xlsx/poi/WithChart.xlsx`
- Measured `inspectWorkbookOpenPlan(bytes, { intent: "edit-plan" })` against `AscendWorkbook.open(bytes, { mode: "full" })`.
- Probed malformed bytes with `inspectWorkbookOpenPlan(new TextEncoder().encode("not a zip"))`.
- Created a synthesis handoff for the product/performance loop instead of adding a new narrow production surface.

## Results

| Fixture | Mode | Review before hydration | Risk families | Parts | Relationships | Median open-plan ms | Median full-open ms | Full/open-plan ratio |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| clean `SampleSS.xlsx` | formula | false | none | 13 | 10 | 0.156 | 2.202 | 14.07x |
| macro `vba.xlsm` | metadata-only | true | preservedMacro | 12 | 9 | 0.092 | 1.408 | 15.27x |
| pivot `ExcelPivotTableSample.xlsx` | formula | false | none | 27 | 19 | 0.189 | 2.235 | 11.83x |
| ActiveX `activex_checkbox.xlsx` | metadata-only | true | preservedActiveX | 17 | 12 | 0.109 | 1.393 | 12.81x |
| chart `WithChart.xlsx` | formula | false | none | 15 | 10 | 0.078 | 1.126 | 14.43x |

Malformed bytes did not produce a mode recommendation; they failed with `Missing end of central directory record`. The proof bundle should represent this as a structured rejection boundary.

## Confidence

High that no additional open-plan surface is needed before the proof bundle. Medium that the benchmark evidence is stable enough for public claims because it is local and small-fixture based; the product loop should rerun it in a repeatable harness before publishing.

## Fold-in decision

Promote to topic synthesis and product/performance handoff. Do not fold into production yet. The implementation owner should build a report/golden proof over existing surfaces rather than a new command.

## Next question

Can token-bounded agent view become the next product claim with a compact proof ladder over omission counters, estimated tokens, and recovery paths across SDK/CLI/API/MCP?
