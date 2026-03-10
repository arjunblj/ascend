# Ascend Agent Instructions

## What This Is

Agent-native spreadsheet platform. TypeScript monorepo with Bun.

## Repo Layout

- `packages/schema` -- shared types, error codes, operation definitions
- `packages/core` -- canonical workbook model (Workbook, Sheet, Cell, SparseGrid, StyleRegistry)
- `packages/formulas` -- formula parser, AST, function implementations
- `packages/engine` -- dependency graph, recalculation, operation/patch engine, diff
- `packages/io-xlsx` -- XLSX/XLSM reader and writer with preservation capsules
- `packages/io-csv` -- CSV/TSV reader and writer
- `packages/verify` -- structural checker, formula linter, dependency tracer
- `packages/sdk` -- public SDK wrapping everything into a Workbook API
- `apps/cli` -- CLI binary
- `apps/api` -- HTTP API server
- `apps/mcp` -- MCP tool server
- `fixtures/` -- test workbooks, formula conformance cases, golden snapshots

## Commands

```bash
bun install                    # install deps
bun test --recursive           # run all tests
bunx biome check               # lint + format check
bunx biome check --write       # auto-fix
bunx tsc --build               # typecheck all packages
```

## Conventions

- Conventional commits: `feat(scope):`, `fix(scope):`, `test(scope):`, `chore:`, `ci:`, `docs:`
- Scopes match package names: `schema`, `core`, `formulas`, `engine`, `io-xlsx`, `io-csv`, `verify`, `sdk`, `cli`, `api`, `mcp`
- Plain TypeScript everywhere. No frameworks in the engine packages.
- Prefer one production implementation over layered legacy/compatibility paths when changing behavior.
- Self-documenting code by default. Comments should be rare and only used when they clarify nuanced semantics, tricky edge cases, or non-obvious invariants that a human maintainer could otherwise misread.
- All engine functions are pure: `(input) => Result<output, error>`. No side effects, no ambient state.
- Tests encode behavior and edge cases, not implementation details.
- Keep user-facing formula/read metadata symbolic and explainable; do not eagerly flatten advanced references if that would hide what the workbook actually contains.
- For substantive work, keep code quality, performance, and maintainability moving together: prefer simpler designs, shared helpers, and validation via tests, typecheck, lint, and benchmarks.

## Package Dependencies

schema < core < formulas < engine < verify < sdk
schema < core < io-xlsx
schema < core < io-csv
sdk < cli, api, mcp
