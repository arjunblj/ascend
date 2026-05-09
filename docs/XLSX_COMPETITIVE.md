# XLSX / spreadsheet stack comparison (high level)

| Area | Ascend | SheetJS (`xlsx`) | ExcelJS | HyperFormula |
|------|--------|------------------|---------|--------------|
| XLSX read/write | Yes, preservation capsules | Yes | Yes | No (values feed) |
| Formula evaluation | Yes (engine + graph) | No | No | Yes |
| License | Apache-2.0 | See SheetJS terms | MIT | GPLv3 (commercial license available) |
| Agent surfaces | SDK, CLI, MCP, HTTP | Library only | Library only | Library only |

Ascend optimizes for **one TS-native pipeline**: IO → model → formulas → incremental recalc → verify → SDK, with **byte-level preservation** for unknown OOXML parts on full loads.

For deeper research notes see [SPREADSHEET_ENGINE_LANDSCAPE.md](./SPREADSHEET_ENGINE_LANDSCAPE.md) and [xlsx-compatibility-research.md](./xlsx-compatibility-research.md).
