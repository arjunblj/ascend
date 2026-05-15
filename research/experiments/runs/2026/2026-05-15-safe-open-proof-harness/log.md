# Safe Open Proof Harness

## Question

Can Ascend turn the safe unknown workbook opening proof bundle into a repeatable tracked report generator without adding another SDK, CLI, API, or MCP product surface?

## Hypothesis

Yes. A benchmark/proof harness under `fixtures/benchmarks` can reuse existing `inspectWorkbookOpenPlan()` behavior, public workbook fixtures, and code-generated package edge cases to produce Markdown/JSON proof output while keeping the user-facing surface unchanged.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/
- DuckDB Excel extension / `read_xlsx`: https://duckdb.org/docs/lts/core_extensions/excel.html
- openpyxl tutorial and preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html

## Why this matters to Ascend

The release claim board ranks "safe unknown workbook opening" as the top product/performance proof. The previous proof lived in an ignored probe and synthesis. A tracked harness makes the proof repeatable without promoting a new product surface or making stronger claims than the implementation can support.

## Probe/implementation

- Inspected `packages/sdk/src/open-plan.ts`, open-plan tests, prior safe-open proof logs, and fixture inventory.
- Added `fixtures/benchmarks/safe-open-proof.ts`.
  - Uses existing `inspectWorkbookOpenPlan()` and `AscendWorkbook.open(..., { mode: "full" })`.
  - Covers six public fixture files: clean, formula-heavy, macro, pivot, ActiveX, and chart.
  - Generates durable synthetic signed and unknown-part packages in code.
  - Includes malformed bytes as an explicit rejection boundary.
  - Emits Markdown by default and JSON with `--json`.
- Added `fixtures/benchmarks/safe-open-proof.test.ts`.
  - Asserts case coverage.
  - Asserts routing decisions without timing thresholds.
  - Asserts report wording includes honest boundaries.
- Updated `research/experiments/syntheses/2026-05-safe-open-proof-bundle.md` with the tracked harness and fresh run output.

Local proof command:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1
```

## Results

| Case | Fixture | Bytes | Mode | Review before hydration | Risk families | Parts | Relationships | Median open-plan ms | Median full-open ms | Full/open-plan ratio | Boundary |
| --- | --- | ---: | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| clean | `fixtures/xlsx/poi/SampleSS.xlsx` | 9112 | formula | false | none | 13 | 10 | 0.186 | 1.978 | 10.64x | ok |
| formula-heavy | `fixtures/xlsx/poi/formula_stress_test.xlsx` | 64769 | formula | false | none | 27 | 22 | 0.197 | 6.514 | 33.05x | ok |
| macro | `fixtures/xlsx/calamine/vba.xlsm` | 12752 | metadata-only | true | preservedMacro | 12 | 9 | 0.073 | 1.602 | 21.93x | ok |
| pivot | `fixtures/xlsx/poi/ExcelPivotTableSample.xlsx` | 19460 | formula | false | none | 27 | 19 | 0.143 | 2.143 | 15.00x | ok |
| ActiveX | `fixtures/xlsx/libreoffice/activex_checkbox.xlsx` | 12433 | metadata-only | true | preservedActiveX | 17 | 12 | 0.096 | 1.595 | 16.62x | ok |
| chart | `fixtures/xlsx/poi/WithChart.xlsx` | 10138 | formula | false | none | 15 | 10 | 0.090 | 1.355 | 15.11x | ok |
| signed | synthetic digital-signature package | 2254 | metadata-only | true | preservedSignature | 8 | 4 | 0.055 | 0.091 | 1.66x | ok |
| unknown part | synthetic unknown package part | 1697 | metadata-only | true | preservedOther | 6 | 3 | 0.036 | 0.081 | 2.26x | ok |
| malformed | synthetic malformed bytes | 9 | rejected | n/a | n/a | n/a | n/a | n/a | n/a | n/a | open-plan rejected: Missing end of central directory record |

Validation passed:

- `bun test fixtures/benchmarks/safe-open-proof.test.ts`
- `bun test packages/sdk/src/open-plan.test.ts`
- `bun test apps/cli/src/cli.test.ts -t "open-plan"`
- `bun test apps/api/api.test.ts -t "open-plan"`
- `bun test apps/mcp/src/index.test.ts -t "open_plan"`
- `bunx biome check fixtures/benchmarks/safe-open-proof.ts fixtures/benchmarks/safe-open-proof.test.ts`

Validation caveat:

- `bunx tsc --build` is blocked by unrelated dirty changes in `packages/io-xlsx/src/reader/sheet.ts` (`rowIndexAttr` argument mismatch/name error). This cycle did not touch that file.

## Confidence

High that the harness closes the repeatability gap for the claim proof. Medium that the timing numbers are publishable as-is; they are local measurements and should be rerun in the release environment before publication.

## Fold-in decision

Folded into the performance/product proof loop as a tracked benchmark/proof harness. No new product surface was added.

## Next question

Can auditable package-part mutation be reduced to one stable proof schema over passthrough, regenerate, add, drop, and error without changing writer behavior?
