# Safe open proof bundle

## Question

Can the safe unknown workbook opening proof bundle be generated from existing SDK/CLI/API/MCP open-plan surfaces over public fixtures, with latency evidence and no new product surface?

## Hypothesis

Yes. `inspectWorkbookOpenPlan()` already owns the package-level routing contract. The missing product proof is a fixture-backed report that shows recommendations, review boundaries, risk families, package counts, latency contrast, and malformed-package rejection.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Excel digital signatures and code signing: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing
- Microsoft active content types: https://support.microsoft.com/en-gb/office/active-content-types-in-your-files-b7ff2e8a-4055-47d4-8c7d-541e19f62bea
- openpyxl tutorial and preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

The claim map ranked "safe unknown workbook opening" first because it is already product-shaped and can be proven without touching writer correctness. Agents need a cheap, explainable first action before they hydrate unknown workbooks or plan edits around active content, signatures, unknown parts, or expensive metadata.

## Probe/implementation

- Inspected current open-plan implementation and tests:
  - `packages/sdk/src/open-plan.ts`
  - `packages/sdk/src/open-plan.test.ts`
  - CLI/API/MCP open-plan routing
  - prior safe-open proof syntheses and logs
- Created a local ignored probe script under `research/experiments/runs/2026/2026-05-15-safe-open-proof-bundle/probes/`.
- Ran the probe with 9 repeated samples after warmup against:
  - public clean, formula-heavy, macro, pivot, ActiveX, and chart fixtures;
  - synthetic digital-signature and unknown-package-part workbooks;
  - synthetic malformed bytes.
- Updated `research/experiments/syntheses/2026-05-safe-open-proof-bundle.md` with the proof bundle, boundary language, fresh timings, and next handoff.
- Did not change production code or add another product surface.

## Results

| Case | Fixture | Bytes | Mode | Review before hydration | Risk families | Parts | Relationships | Median open-plan ms | Median full-open ms | Full/open-plan ratio | Boundary |
| --- | --- | ---: | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| clean | `fixtures/xlsx/poi/SampleSS.xlsx` | 9112 | formula | false | none | 13 | 10 | 0.114 | 1.816 | 15.87x | ok |
| formula-heavy | `fixtures/xlsx/poi/formula_stress_test.xlsx` | 64769 | formula | false | none | 27 | 22 | 0.178 | 5.875 | 33.04x | ok |
| macro | `fixtures/xlsx/calamine/vba.xlsm` | 12752 | metadata-only | true | preservedMacro | 12 | 9 | 0.058 | 1.209 | 20.73x | ok |
| pivot | `fixtures/xlsx/poi/ExcelPivotTableSample.xlsx` | 19460 | formula | false | none | 27 | 19 | 0.142 | 2.027 | 14.31x | ok |
| ActiveX | `fixtures/xlsx/libreoffice/activex_checkbox.xlsx` | 12433 | metadata-only | true | preservedActiveX | 17 | 12 | 0.091 | 1.581 | 17.41x | ok |
| chart | `fixtures/xlsx/poi/WithChart.xlsx` | 10138 | formula | false | none | 15 | 10 | 0.068 | 1.097 | 16.09x | ok |
| signed | synthetic digital-signature package | 2293 | metadata-only | true | preservedSignature | 8 | 4 | 0.040 | 0.090 | 2.25x | ok |
| unknown part | synthetic unknown package part | 1767 | metadata-only | true | preservedOther | 6 | 3 | 0.037 | 0.065 | 1.75x | ok |
| malformed | synthetic malformed bytes | 9 | rejected | n/a | n/a | n/a | n/a | n/a | n/a | n/a | open-plan rejected: Missing end of central directory record |

The product proof should use this language:

- Open-plan is pre-hydration package routing.
- `metadata-only` plus `reviewBeforeHydration: true` is the correct branch for active content, security material, and unknown package parts.
- Malformed bytes are a rejection boundary, not a safe-mode recommendation.
- The timing probe supports direction, not a CI threshold.

## Confidence

High that no additional SDK/CLI/API/MCP surface is needed. Medium-high that the proof bundle is ready for product/performance follow-up, because the public fixture cases are real and repeatable, while the signed and unknown-part cases should become durable fixtures before a published claim.

## Fold-in decision

Promote to product/performance as a proof bundle over existing open-plan surfaces. Do not fold production code in this cycle.

## Next question

Can auditable package-part mutation be reduced to one stable proof schema over passthrough, regenerate, add, drop, and error, without changing writer behavior?
