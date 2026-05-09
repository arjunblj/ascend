# Ascend

Agent-native spreadsheet platform. Read, edit, verify, and export `.xlsx` workbooks from TypeScript with a dependency graph, incremental recalc, and preservation-oriented XLSX I/O.

## Requirements

- [Bun](https://bun.sh) 1.3+ (recommended: latest stable)

## Install (from this repo)

```bash
git clone <repo-url> ascend && cd ascend
bun install
```

Use the workspace packages from your app (path dependency) or `bun link` inside `packages/sdk` until npm packages ship.

```json
{
  "dependencies": {
    "@ascend/sdk": "workspace:*"
  }
}
```

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

## API overview

| Surface | Role |
|--------|------|
| **`Ascend`** | Shorthand: `open`, `create`, `fromCsv`, `listOperations`, `getOperationsSchema` |
| **`AscendWorkbook`** | Mutable workbook: `apply`, `recalc`, `check`, `lint`, `save`, `sheet`, streaming reads |
| **`WorkbookDocument`** | Read-heavy, cacheable view over bytes (inspect without full mutation path) |
| **`ops`** | Typed operation builders aligned with the patch engine |
| **`streamWorkbookRows`** | Row-oriented streaming for large sheets |

Generate HTML API reference: `bun run docs:sdk` (output: `dist/docs/sdk-api`). OpenAPI for the HTTP server: [docs/openapi.yaml](docs/openapi.yaml).

## Performance (CI targets)

Synthetic benchmark gates in [`fixtures/benchmarks/targets.ts`](fixtures/benchmarks/targets.ts) (best scenario per category):

| Category | Minimum throughput |
|----------|-------------------|
| Read | 3M cells/s |
| Write | 1.5M cells/s |
| Recalc | 500K cells/s |

Run locally: `bun bench`, or `bun bench --scenario <name> --repeat 5 --json`. See [BENCHMARKS.md](BENCHMARKS.md) for methodology and competitive comparisons.

## HTTP API

Start the dev server from `apps/api`:

```bash
bun run apps/api/src/index.ts
```

Endpoints (JSON bodies, local file paths): `GET /health`, `POST /inspect`, `/read`, `/agent-view`, `/write`, `/preview`, `/calc`, `/check`, `/lint`, `/trace`, `/diff`, `/export`. See [docs/openapi.yaml](docs/openapi.yaml).

## MCP (Cursor / Claude Desktop)

Run the MCP server (stdio):

```bash
bun run apps/mcp/src/index.ts
```

**Cursor** — add to MCP config (e.g. `~/.cursor/mcp.json`):

```json
{
  "mcpServers": {
    "ascend": {
      "command": "bun",
      "args": ["run", "/absolute/path/to/ascend/apps/mcp/src/index.ts"],
      "cwd": "/absolute/path/to/ascend"
    }
  }
}
```

**Claude Desktop** — in `claude_desktop_config.json` under `mcpServers`, use the same `command` / `args` / `cwd` pattern pointing at your clone.

Tools exposed: `ascend.inspect`, `ascend.list_sheets`, `ascend.read`, `ascend.find`, `ascend.agent_view`, `ascend.preview`, `ascend.write`, `ascend.calc`, `ascend.list_operations`, `ascend.check`, `ascend.lint`, `ascend.trace`, `ascend.diff`, `ascend.export`.

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

- `packages/schema` — shared types, error codes, operation definitions
- `packages/core` — canonical workbook model
- `packages/formulas` — formula parser, AST, **~450+ registered function names** (Excel aliases included; not every Excel function)
- `packages/engine` — dependency graph, deterministic recalculation, operation/patch engine
- `packages/io-xlsx` — XLSX/XLSM reader and writer with preservation capsules
- `packages/io-csv` — CSV/TSV reader and writer
- `packages/verify` — structural checker, formula linter, dependency tracer
- `packages/sdk` — public SDK wrapping everything
- `apps/cli` — CLI binary
- `apps/api` — HTTP API server
- `apps/mcp` — MCP tool server for AI agents

## Examples

See [`examples/`](examples/) for runnable scripts (`read-modify-save`, `create-from-scratch`, CSV, batch ops, MCP notes).

## Security note

Formula evaluation may use **code generation** (`new Function`) for hot paths. Only evaluate workbooks and formulas from **trusted** sources unless you run in an isolated process. See [docs/SECURITY.md](docs/SECURITY.md).

## Development

```bash
bun install
bun test --recursive
bunx biome check
bunx tsc --build
```

## License

Apache-2.0
