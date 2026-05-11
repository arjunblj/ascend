# Ascend

Agent-native spreadsheet automation for TypeScript.

Read, edit, recalculate, verify, and export real `.xlsx` workbooks without collapsing them into raw cell dumps.

Ascend focuses on four surfaces together:

- XLSX/XLSM read and write with package-level preservation for workbook features that are not fully modeled yet.
- Deterministic formulas, dependency tracing, structural operations, and workbook verification.
- SDK, CLI, HTTP, and MCP interfaces for automation and agent workflows.
- Reproducible corpus tests and benchmarks that report correctness, throughput, and memory evidence.

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

All commands support machine-readable output where it is useful for automation.

## Agent Surfaces

- HTTP API: start `apps/api/src/index.ts`; OpenAPI lives at [docs/openapi.yaml](docs/openapi.yaml).
- MCP server: run `apps/mcp/src/index.ts` for `ascend.inspect`, `ascend.read`, `ascend.write`, `ascend.calc`, `ascend.check`, `ascend.trace`, and related tools.
- Operation schemas: `ascend ops --json` exposes typed operations and recovery guidance.

## Benchmarks

```bash
bun bench --repeat 5 --json
bun run bench:competitive-io --workload all --repeat 5 --json
bun run bench:formula:sota --profile all --repeat 5 --json
```

Benchmarks emit machine-readable JSON and include correctness and memory checks where applicable. See [BENCHMARKS.md](BENCHMARKS.md) for methodology.

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

Apache-2.0
