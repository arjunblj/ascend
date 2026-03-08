# Ascend

Agent-native spreadsheet platform. Read, edit, verify, and export `.xlsx` workbooks from TypeScript.

## Quickstart

```typescript
import { Ascend } from '@ascend/sdk'

const wb = await Ascend.open('model.xlsx')

const sheet = wb.sheet('Revenue')
const data = sheet?.range('A1:D10')

wb.apply([
  { op: 'setCells', sheet: 'Revenue', updates: [{ ref: 'B2', value: 500 }] },
  { op: 'setFormula', sheet: 'Revenue', ref: 'E2', formula: '=B2*C2' },
])

wb.recalc()
wb.check()
wb.lint()

await wb.save('output.xlsx')
```

## CLI

```bash
ascend create output.xlsx
ascend inspect model.xlsx
ascend read model.xlsx A1:D10 --sheet Revenue
ascend write model.xlsx A1 '[100, 200, 300]'
ascend calc model.xlsx
ascend check model.xlsx
ascend lint model.xlsx
ascend trace model.xlsx 'Revenue!E2'
ascend diff before.xlsx after.xlsx
ascend export model.xlsx output.csv --format csv
```

All commands support `--json` for machine-readable output.

## Architecture

- `packages/schema` -- shared types, error codes, operation definitions
- `packages/core` -- canonical workbook model
- `packages/formulas` -- formula parser + ~109 Excel-compatible functions
- `packages/engine` -- dependency graph, deterministic recalculation, operation/patch engine
- `packages/io-xlsx` -- XLSX/XLSM reader and writer with preservation capsules
- `packages/io-csv` -- CSV/TSV reader and writer
- `packages/verify` -- structural checker, formula linter, dependency tracer
- `packages/sdk` -- public SDK wrapping everything
- `apps/cli` -- CLI binary
- `apps/api` -- HTTP API server
- `apps/mcp` -- MCP tool server for AI agents

## Development

```bash
bun install
bun test --recursive
bunx biome check
bunx tsc --build
```

## License

Apache-2.0
