# Feature Fingerprint Open Mode

## Question

Can a workbook feature fingerprint choose the ideal open mode before full hydration?

## Hypothesis

Yes. A package-level fingerprint can cheaply identify enough workbook risk and workload shape to choose between `metadata-only`, `values`, `formula`, and `formula + richMetadata` before hydrating all cells and rich sheet metadata.

## External sources checked

- Microsoft OPC fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- ECMA-376 standard page, especially Part 2 Open Packaging Conventions: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- SheetJS parse options, including `bookProps`, `bookSheets`, `sheetRows`, `cellFormula`, `bookFiles`, `bookVBA`: https://docs.sheetjs.com/docs/api/parse-options
- openpyxl tutorial `load_workbook` flags for `read_only`, `data_only`, `keep_vba`, `keep_links`, and shape-loss warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- openpyxl optimized modes for large read-only workbooks: https://openpyxl.pages.heptapod.net/openpyxl/optimized.html
- calamine Reader trait and metadata-oriented API surface: https://docs.rs/calamine/latest/calamine/trait.Reader.html
- ExcelJS README/npm read options, including `ignoreNodes` and streaming reader: https://www.npmjs.com/package/exceljs

## Why this matters to Ascend

Ascend already has richer open modes than most OSS libraries, but callers still have to know which mode to choose. A first-class open planner would make agent reads faster, reduce accidental partial-load edits, and surface macro/signature/pivot/chart/slicer risks before an agent asks to mutate a workbook.

## Probe/implementation

Inspected local implementation:

- `packages/io-xlsx/src/reader/index.ts` exposes `mode`, `sheets`, `maxRows`, `richMetadata`, and partial-load metadata.
- `packages/io-xlsx/src/package-graph.ts` already classifies package parts into feature families and preservation policies.
- `packages/sdk/src/load.ts` strips source archives for partial loads, correctly preventing partial workbooks from pretending to be write-ready.
- API/MCP routes already use mode choices manually, for example `metadata-only` for package/feature inventory and `formula` for agent views.

Ignored probes:

- `probes/feature-fingerprint.ts` classified six local workbooks by part families, formula tags, cell tags, and preservation policies.
- `probes/open-mode-timing.ts` compared median local timings for package fingerprinting and existing open modes.

Timing sample, median of five warm runs:

| File | Fingerprint | Metadata | Values | Formula | Formula + Rich |
| --- | ---: | ---: | ---: | ---: | ---: |
| `conditional-formatting.xlsx` | 0.377 ms | 1.593 ms | 5.698 ms | 2.221 ms | 4.913 ms |
| `ms-excel-formulas-and-pivot-tables.xlsx` | 1.095 ms | 2.192 ms | 151.086 ms | 478.582 ms | 618.358 ms |
| `excel-dashboard-v2.xlsx` | 0.722 ms | 2.375 ms | 1551.698 ms | 1083.964 ms | 1871.745 ms |
| `bevreport-demo.xlsm` | 8.294 ms | 9.8 ms | 134.772 ms | 393.704 ms | 364.603 ms |

## Results

The feature signal is strong enough to drive an open recommendation:

- Macro/ActiveX workbooks can be routed to `metadata-only` risk inventory before edit-capable load.
- Pivot/chart/slicer/calc-chain/formula workbooks should prefer `formula + richMetadata` before write planning, but may still use `metadata-only` for first risk inventory.
- Simple value workbooks can use `values` or `metadata-only` depending on caller intent.
- Fingerprinting is much cheaper than hydrating cells on medium formula/dashboard workbooks.

Competitor contrast:

- SheetJS exposes `bookProps`/`bookSheets` and sheet row limits, but the docs frame these as caller-selected parse options rather than a preservation-aware planner.
- openpyxl exposes read-only/data-only/keep-vba switches but warns that not all Excel items are read and shapes may be lost on save.
- ExcelJS exposes ignored XML nodes and streaming reader options, but not a package-risk open planner.
- calamine is read-focused and exposes metadata methods, but not edit-preservation planning.

## Confidence

High that this is worth folding into product/DX and performance planning. Medium on exact thresholds because the probe used local files and in-process warm timings, not benchmark-isolated measurements.

## Fold-in decision

Promote to product/DX and performance loops. Do not fold directly into production from this cycle. The next implementation slice should be an internal `inspectOpenPlan(bytes, intent)` or SDK method that returns recommended mode, reasons, risk features, and estimated cost class, with tests over small fixture bytes.

## Next question

Can a production-safe open planner be implemented from the existing package graph without inflating worksheet bodies, and can it recommend `metadata-only` for active-content workbooks while recommending `formula + richMetadata` for formula/pivot dashboards?
