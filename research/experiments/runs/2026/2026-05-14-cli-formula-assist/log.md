# CLI Formula Assist

## Question

Should Ascend expose the existing formula IDE helper through the CLI so agents can get diagnostics, completions, signature help, reference insertion, and F4-style reference cycling before planning a workbook edit?

## Hypothesis

Yes. API, MCP, and SDK already expose the helper, but CLI-only agents had to use HTTP/MCP or hand-roll formula checks. A small CLI subcommand can close that product/DX gap without changing formula semantics or write behavior.

## External sources checked

- Microsoft Formula AutoComplete: https://support.microsoft.com/en-gb/office/use-formula-autocomplete-6d13daa5-e003-4431-abab-9edef51fae6b
- Language Server Protocol 3.18 specification: https://github.com/microsoft/language-server-protocol/blob/gh-pages/_specifications/lsp/3.18/specification.md
- HyperFormula named expressions guide: https://hyperformula.handsontable.com/guide/named-expressions.html
- HyperFormula API documentation: https://hyperformula.handsontable.com/docs/api/classes/hyperformula.html

## Why this matters to Ascend

Formula edits are one of the easiest places for agents to make plausible but invalid workbook changes. Excel-style autocomplete and LSP-style diagnostics/signature help both point toward interactive assistance before mutation. Ascend already has formula spans and formula-assist internals; exposing them through CLI gives headless agents a no-server path to repair formulas before `plan`.

## Probe/implementation

- Inspected `apps/cli/src/commands/formula.ts`: CLI formula surface only supported `show`, `set`, and `fill`.
- Inspected `packages/sdk/src/formula-edit.ts`, `apps/api/src/server.ts`, and `apps/mcp/src/index.ts`: `formulaAssist` already returns diagnostics, tokens, active reference, hover, completions, signature, signature help, code actions, insertion preview, and reference cycling.
- Added `ascend formula assist '<formula>'` with flags:
  - `--cursor`
  - `--prefix`
  - `--completion-limit`
  - `--function-name`
  - `--reference`
  - `--replace-reference-at-cursor`
  - `--cycle-reference`
  - `--json`
- Added a CLI test proving `formula assist --json` works without opening a workbook and returns parse diagnostics, SUM completion/signature help, insertion preview, and reference cycling.
- Updated agent-init, `docs/AGENT_API.md`, `llms.txt`, and `llms-full.txt` so agents discover the CLI route alongside API/MCP/SDK routes.

## Results

- `bun test apps/cli/src/cli.test.ts -t "formula assist"` passed.
- `bun test apps/cli/src/cli.test.ts -t "formula"` passed.
- `bun test apps/cli/src/cli.test.ts -t "agent-init"` passed.
- `bunx biome check apps/cli/src/commands/formula.ts apps/cli/src/index.ts apps/cli/src/cli.test.ts apps/cli/src/commands/agent-init.ts docs/AGENT_API.md llms.txt llms-full.txt research/experiments/index.md research/experiments/runs/2026/2026-05-14-cli-formula-assist/log.md` passed.
- `bunx tsc --build` passed.
- `bun run test:changed` passed with 4966 tests, 1 skip, and 0 failures.

## Confidence

High. This is a thin CLI wrapper over an existing tested SDK helper, with no workbook hydration, no writes, and no formula semantics changes.

## Fold-in decision

Folded into the product/DX loop as a CLI command and agent documentation update.

## Next question

Can Ascend expose package-action proof summaries through CLI/API/MCP release-proof workflows so agents can see passthrough/regenerate/add/drop/error evidence without reading the full SDK object?
