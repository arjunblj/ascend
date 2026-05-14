# Open Plan CLI Explain First

## Question

Can CLI users and agents access the workbook open planner as an explain-first recommendation without changing existing read, inspect, or commit behavior?

## Hypothesis

Yes. A dedicated CLI command can expose the planner result as JSON or concise text, letting agents choose a load mode from package features while preserving existing command defaults.

## External sources checked

- Model Context Protocol tool documentation on named tools with schemas and descriptions: https://modelcontextprotocol.io/docs/concepts/tools
- Notion MCP overview as a current example of agent-facing tools exposed through explicit capabilities: https://developers.notion.com/guides/mcp/overview
- Notion developer platform page showing CLI, MCP, API, SDK, and workers as parallel agent surfaces: https://www.notion.com/product/dev
- SheetJS parse options for caller-selected workbook metadata, sheet, raw file, formula, row-limit, and VBA options: https://docs.sheetjs.com/docs/api/parse-options/
- openpyxl tutorial `load_workbook` options and warning that not all Excel items are read/preserved: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- openpyxl optimized read-only mode for large files: https://openpyxl.pages.heptapod.net/openpyxl/optimized.html

## Why this matters to Ascend

The SDK open planner is useful only if agents can discover and call it before hydration. A CLI command is the narrowest product/DX fold-in: it gives local agents a stable, inspectable contract while avoiding hidden behavior changes in established commands.

## Probe/implementation

Local inspection:

- `apps/cli/src/index.ts` centralizes command registration, help text, allowed flags, and closest-command hints.
- Existing commands choose load modes manually; `inspect --detail package-graph` and `inspect --agent` expose adjacent evidence but require different command semantics.
- `apps/cli/src/cli.test.ts` already covers structured JSON envelopes and active-content fixtures.

Implementation:

- Added `apps/cli/src/commands/open-plan.ts`.
- Registered `ascend open-plan <file> [--intent ...] [--json]`.
- The command reads bytes, calls `inspectWorkbookOpenPlan`, and prints either the full machine JSON plan or a concise text summary with reasons and risk features.
- Existing command defaults are unchanged.

Validation:

- `bun test apps/cli/src/cli.test.ts -t "open-plan"` passed.
- `bun test apps/cli/src/cli.test.ts -t "--help"` passed.
- `bunx biome check apps/cli/src/commands/open-plan.ts apps/cli/src/index.ts apps/cli/src/cli.test.ts` passed.
- `bunx tsc --build` passed.
- `bun run test:changed` ran the full suite and passed: 4951 pass, 1 skip, 0 fail.

## Results

The CLI fold-in is small and useful:

- `--intent read-values --json` returns `{ mode: "values" }` for a simple workbook.
- Macro packages are routed to `metadata-only` with `reviewBeforeHydration: true`.
- The command is explicit, so it does not silently change `read`, `inspect`, `agent-view`, or `commit`.
- JSON output mirrors the SDK result, which keeps CLI, MCP, and API follow-up work aligned around one schema.

## Confidence

High that the CLI surface is worthwhile and low-risk. Medium that this should become an MCP/API tool unchanged; those surfaces need their own schema and prompt-description checks.

## Fold-in decision

Folded into production as `ascend open-plan`. Promote the same schema to MCP/API only after one more experiment verifies whether tool descriptions and endpoint names make agents choose it before expensive hydration.

## Next question

Can MCP expose `ascend.open_plan` with a description that reliably steers agents to plan first for unknown workbooks, without bloating the MCP tool list or duplicating package-graph output?
