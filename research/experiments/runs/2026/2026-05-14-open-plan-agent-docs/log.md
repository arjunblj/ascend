# Open Plan Agent Docs

## Question

Should `agent-init` and bundled agent docs teach the open-plan -> inspect/read/agent-view -> plan/commit route as the default unknown-workbook workflow?

## Hypothesis

Yes. The new SDK, CLI, MCP, and API open-plan surfaces only become useful if agent-facing docs put them before hydration while still preserving the trust preflight before reading workbook text.

## External sources checked

- MCP tools documentation, especially model-controlled tool selection from descriptions: https://modelcontextprotocol.io/docs/concepts/tools
- Cloudflare Agents MCP tools docs on descriptions and schemas being used by LLMs to decide when to call a tool: https://developers.cloudflare.com/agents/model-context-protocol/tools/
- Google Sheets `spreadsheets.get` guidance to retrieve only specific fields for large spreadsheets: https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/get
- Microsoft Graph Excel session guidance for making workbook access mode explicit: https://learn.microsoft.com/en-us/graph/excel-manage-sessions
- "Model Context Protocol (MCP) Tool Descriptions Are Smelly!" on explicit usage guidelines and compact tool descriptions: https://arxiv.org/abs/2602.14878
- "From Docs to Descriptions: Smell-Aware Evaluation of MCP Server Descriptions" on accuracy, completeness, concision, and behavior boundaries in agent tool docs: https://arxiv.org/abs/2602.18914

## Why this matters to Ascend

The open planner is a routing primitive. If docs continue to start at `inspect`, agents will still hydrate or scan workbook content before asking the cheaper package-level question. This weakens the performance and safety value of the prior fold-ins.

## Probe/implementation

Local inspection:

- `agent-init` did not mention `open-plan`.
- `docs/AGENT_WORKFLOW.md`, `docs/AGENT_API.md`, `llms.txt`, and `llms-full.txt` started with trust or inspect.
- HTTP and MCP safe-edit examples did not include `/open-plan` or `ascend.open_plan`.

Implementation:

- Added open-plan as the first workflow step for unknown XLSX/XLSM files in `agent-init`.
- Added `openPlan` to the machine-readable `agent-init --json` command map.
- Updated `AGENT_WORKFLOW`, `AGENT_API`, `llms.txt`, and `llms-full.txt` to teach open-plan before hydration while keeping trust preflight before reading workbook text.
- Updated HTTP and MCP safe-edit examples with `/open-plan` and `ascend.open_plan` calls.
- Updated the CLI test for the `agent-init` contract.

Validation:

- `bun test packages/sdk/src/agent-docs.test.ts` passed.
- `bun test apps/cli/src/cli.test.ts -t "agent-init"` passed.
- `bunx biome check apps/cli/src/commands/agent-init.ts apps/cli/src/cli.test.ts docs/AGENT_WORKFLOW.md docs/AGENT_API.md llms.txt llms-full.txt examples/agent-safe-edit-http.md examples/agent-safe-edit-mcp.md examples/README.md` passed.
- `bunx tsc --build` passed.
- `bun run test:changed` ran the full suite and passed: 4959 pass, 1 skip, 0 fail.

## Results

The docs now describe a safer route:

1. Run open-plan for unknown XLSX/XLSM files.
2. If `reviewBeforeHydration` is true, stay in metadata/trust/package inventory.
3. Run trust preflight before reading workbook text into the prompt.
4. Inspect/read only the required surface, then plan/commit/verify.

## Confidence

High. This is a small docs and command-contract fold-in with targeted tests. The remaining uncertainty is whether agents actually follow the order in practice; that needs a trace-mining or prompt-harness experiment.

## Fold-in decision

Folded into product/DX docs and `agent-init`. Promote a future evaluation harness that checks tool-call ordering on unknown workbook tasks.

## Next question

Can a small agent-workflow prompt harness detect whether agents call open-plan before inspect/read on unknown XLSX/XLSM tasks?
