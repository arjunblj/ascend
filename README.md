# Ascend

Open-source spreadsheet engine for TypeScript apps and agents.

Ascend loads Excel files into a workbook model you can inspect, patch with typed operations, recalculate, verify, and save with package-level preservation for unmodeled OOXML parts.

Features:

- `.xlsx` / `.xlsm` round trips with package-level preservation.
- Formula execution, dependency tracing, structural edits, and workbook checks.
- SDK, CLI, HTTP API, and MCP server for app and agent workflows.
- Corpus tests and benchmark runs that report correctness, throughput, and memory.

## Install

Requires [Bun](https://bun.sh) 1.3+.

```bash
git clone https://github.com/arjunblj/ascend.git
cd ascend
bun install
bun run examples/create-from-scratch.ts ./out.xlsx
```

Use the SDK from another local app until packages ship:

```bash
cd packages/sdk
bun link
cd /path/to/your-app
bun link @ascend/sdk
```

## Quickstart

```typescript
import { Ascend } from '@ascend/sdk'

const wb = Ascend.create()

wb.apply([
  { op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 21 }] },
  { op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1*2' },
])

wb.recalc()
console.log(wb.sheet('Sheet1')?.cell('B1')?.value)

const check = wb.check()
if (!check.valid) throw new Error(check.issues[0]?.message ?? 'Workbook check failed')

await wb.save('output.xlsx')
```

## API

- `Ascend.open(path | bytes)` opens a workbook.
- `Ascend.create()` creates a workbook.
- `wb.sheet(name).range('A1:D10')` reads cells.
- `wb.apply(ops)`, `wb.recalc()`, `wb.check()`, and `wb.save(path)` edit, calculate, verify, and write.

## Excel Coverage

| Layer | Current coverage |
|-------|------------------|
| Formula engine | 496 registered Excel function names across math, stats, financial, lookup/reference, text, date/time, logical, dynamic arrays, database, engineering, info, forecast, and compatibility aliases |
| Calculation model | Dependency graph, incremental recalculation, structured references, defined names, array/spill values, and traceable formula lineage |
| Editable workbook model | Sheets, cells, ranges, formulas, row/column operations, tables, filters, styles, merges, comments, hyperlinks, validations, and conditional formats |
| XLSX/XLSM preservation | Charts, pivots, macros, drawings/images, slicers, external links, workbook relationships, and unknown OOXML parts are retained where they are not fully modeled |
| Verification | Excel cached-value tests, fixture corpora, round-trip preservation checks, formula linting, structural checks, and correctness-gated benchmarks |

## CLI

```bash
ascend inspect model.xlsx --json
ascend read model.xlsx A1:D10 --sheet Revenue
ascend write model.xlsx A1 '[100, 200, 300]'
ascend calc model.xlsx
ascend check model.xlsx
ascend trace model.xlsx 'Revenue!E2'
ascend export model.xlsx output.tsv --format tsv
```

Commands support machine-readable output where it helps automation.

## Agent Surfaces

- HTTP API: start `apps/api/src/index.ts`; OpenAPI lives at [docs/openapi.yaml](docs/openapi.yaml).
- MCP server: run `apps/mcp/src/index.ts` for `ascend.inspect`, `ascend.read`, `ascend.write`, `ascend.calc`, `ascend.check`, `ascend.trace`, and related tools.
- Operation schemas: `ascend ops --json` exposes typed operations and recovery hints.

## Benchmarks

Benchmark snapshot: local median, 5 samples, external-process runners, correctness gates enabled. See [BENCHMARKS.md](BENCHMARKS.md) for methodology.

| Lane | Profile | Size | Ascend | Baseline |
|------|---------|------|--------|---------------|
| XLSX read | Calamine NYC 311 real workbook | 1,000,001 rows, 28.1M non-empty cells | 13.91s, 2.02M cells/s | rust-calamine 26.90s |
| XLSX write | Excelize plain text generation | 102,400 x 50, 5.12M cells | 1.88s, 2.72M cells/s | Excelize 2.47s |
| XLSX write | pyopenxlsx bulk write | 50,000 x 20, 1.0M cells | 0.24s, 4.15M cells/s | pyopenxlsx 0.64s |
| Formula recalc | HyperFormula indexed lookup | 5,000 INDEX/MATCH formulas | 20.3ms | HyperFormula 587.0ms |
| Incremental recalc | Indexed lookup value edit | 5,000 INDEX/MATCH formulas | 0.051ms | HyperFormula 582.7ms |

## Development

```bash
bun install
bun test --recursive
bunx biome check
bunx tsc --build
```

Core packages are plain TypeScript: `schema`, `core`, `formulas`, `engine`, `io-xlsx`, `io-csv`, `verify`, and `sdk`.

## More

- [Examples](examples/)
- [Security](docs/SECURITY.md)

## License

MIT
