# 2026-05 Safe Open Proof Bundle Handoff

Date: 2026-05-14

## Claim

Ascend can route unknown XLSX/XLSM files through a package-level open plan before workbook hydration, recommending a load mode and review step from observable package features.

## Proof Bundle Shape

| Required proof | What to show | Current status |
| --- | --- | --- |
| Fixture | Clean workbook, macro workbook, ActiveX/control workbook, pivot workbook, chart workbook, malformed bytes | Public fixtures exist for all but malformed should be synthetic and checked as rejection |
| Benchmark | Open-plan latency versus full hydration on warmed public fixtures | Local probe shows 11.83x-15.27x faster on five fixtures |
| API/CLI/MCP surface | Existing SDK `inspectWorkbookOpenPlan`, CLI `ascend open-plan`, API `POST /open-plan`, MCP `ascend.open_plan` | Implemented; do not add another surface |
| Validation gate | Focused open-plan tests, API/CLI/MCP tests, docs ordering test, typecheck, Biome, changed tests | Existing tests cover behavior; proof bundle needs one report/golden |
| Competitor contrast | openpyxl warns unsupported items can be lost; SheetJS/ExcelJS expose read/write options; LibreOffice Safe Mode is application recovery rather than pre-hydration workbook routing | Good enough for product contrast if phrased narrowly |
| Honest boundary | Not malware scanning, not a sandbox, not Excel trust-center replacement, not guaranteed inspection for malformed ZIP packages | Must be included in every claim and report |

## Local Probe

| Fixture | Mode | Review before hydration | Risk families | Parts | Relationships | Median open-plan ms | Median full-open ms | Ratio |
| --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| `fixtures/xlsx/poi/SampleSS.xlsx` | formula | false | none | 13 | 10 | 0.156 | 2.202 | 14.07x |
| `fixtures/xlsx/calamine/vba.xlsm` | metadata-only | true | preservedMacro | 12 | 9 | 0.092 | 1.408 | 15.27x |
| `fixtures/xlsx/poi/ExcelPivotTableSample.xlsx` | formula | false | none | 27 | 19 | 0.189 | 2.235 | 11.83x |
| `fixtures/xlsx/libreoffice/activex_checkbox.xlsx` | metadata-only | true | preservedActiveX | 17 | 12 | 0.109 | 1.393 | 12.81x |
| `fixtures/xlsx/poi/WithChart.xlsx` | formula | false | none | 15 | 10 | 0.078 | 1.126 | 14.43x |

Malformed bytes currently fail with `Missing end of central directory record`. That should be reported as a structured rejection, not as a safe mode recommendation.

## Product Handoff

```text
/goal Build the safe unknown workbook opening proof bundle without adding a new product surface. Use existing SDK/CLI/API/MCP open-plan contracts to generate a fixture-backed report over clean, macro, ActiveX/control, pivot, chart, and malformed public cases. Include recommendation, reviewBeforeHydration, risk feature families, package counts, open-plan latency versus full hydration, and structured malformed-package rejection. Validate focused open-plan tests, docs ordering, bunx tsc --build, bunx biome check, and bun run test:changed. Keep boundaries explicit: pre-hydration package routing, not malware scanning or sandboxing.
```

## Do Not Promote Yet

- Do not claim malicious workbook safety.
- Do not claim full Excel compatibility from an open plan.
- Do not add a second CLI/API/MCP surface for the same recommendation.
- Do not publish private corpus measurements.
- Do not turn local latency numbers into thresholds until a stable benchmark harness owns them.
