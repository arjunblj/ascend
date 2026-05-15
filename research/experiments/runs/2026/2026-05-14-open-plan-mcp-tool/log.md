# Open Plan MCP Tool

## Question

Can MCP expose `ascend.open_plan` with a description that reliably steers agents to plan first for unknown workbooks, without bloating the tool list or duplicating package-graph output?

## Hypothesis

Yes. A compact MCP tool that returns the SDK open-plan schema can sit before `inspect`, `read`, `agent_view`, and `plan`, while leaving detailed package evidence in `ascend.package_graph`.

## External sources checked

- Model Context Protocol tool documentation on named tools, descriptions, schemas, and model-controlled invocation: https://modelcontextprotocol.io/docs/concepts/tools
- Notion MCP overview for agent-facing tool connection patterns: https://developers.notion.com/guides/mcp/overview
- Notion developer platform page showing CLI, MCP, API, SDK, and workers as parallel agent surfaces: https://www.notion.com/product/dev
- SheetJS parse options for workbook metadata, sheet limits, raw files, formulas, and VBA controls: https://docs.sheetjs.com/docs/api/parse-options/
- openpyxl optimized read-only mode for large workbooks: https://openpyxl.pages.heptapod.net/openpyxl/optimized.html
- "From Docs to Descriptions: Smell-Aware Evaluation of MCP Server Descriptions" for risks in vague or incomplete MCP tool descriptions: https://arxiv.org/abs/2602.18914

## Why this matters to Ascend

MCP is the agent-native surface where expensive or unsafe workbook hydration is most likely to happen accidentally. An explicit open-plan tool lets agents ask a cheap, preservation-aware question first, then pick `read`, `agent_view`, `inspect`, or `plan` with better evidence.

## Probe/implementation

Local inspection:

- `apps/mcp/src/index.ts` registers all MCP tools in one place and already separates lightweight inspection (`inspect`, `list_sheets`, `active_content`, `package_graph`) from mutation planning and commit.
- `apps/mcp/src/index.test.ts` verifies the registered tool set by name and directly calls handlers through MCP server internals.
- `ascend.package_graph` returns detailed OPC evidence, so `ascend.open_plan` should not duplicate full part lists.

Implementation:

- Added `ascend.open_plan` to the MCP server.
- Tool description tells agents to use it first for unknown XLSX/XLSM files when cost, active content, formulas, rich metadata, or preservation risks could affect the next tool.
- Inputs: `file` and optional `intent` enum (`risk-inventory`, `read-values`, `formula-analysis`, `edit-plan`).
- Output: the SDK `inspectWorkbookOpenPlan` result wrapped in the existing MCP success envelope.
- Added focused tests for read-intent value workbooks and macro workbooks, plus updated the registered-tool count.

Validation:

- `bun test apps/mcp/src/index.test.ts -t "open_plan"` passed.
- `bun test apps/mcp/src/index.test.ts -t "all ascend tools"` passed.
- `bunx biome check apps/mcp/src/index.ts apps/mcp/src/index.test.ts` passed after formatting.
- `bunx tsc --build` passed.
- `bun run test:changed` ran the full suite and passed: 4955 pass, 1 skip, 0 fail.

## Results

The tool is a clean fold-in:

- It adds one named MCP tool and reuses the SDK planner schema.
- It avoids hidden behavior changes in existing MCP tools.
- It keeps package-graph detail separate while still surfacing risk families and reasons.
- The tool description is specific about when to use it and what it does not do: no hydration and no mutation.

## Confidence

High that this should stay. Medium that the description alone is enough to steer all agents; future evaluation should mine tool-call traces or build an MCP prompt-harness that checks whether agents call `ascend.open_plan` before expensive reads.

## Fold-in decision

Folded into production as `ascend.open_plan`. Promote the same shape to API after validating whether `/open-plan` should be a standalone endpoint or folded into `/inspect` responses.

## Next question

Can the HTTP API expose the open planner without encouraging clients to skip existing trust reports and package-graph audits?
