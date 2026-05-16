# Ascend examples

Run from the **repository root** with Bun:

```bash
bun run examples/create-from-scratch.ts
bun run examples/read-modify-save.ts ./path/to/file.xlsx
bun run examples/formula-eval.ts
bun run examples/csv-convert.ts
bun run examples/batch-ops.ts
bun run examples/agent-safe-edit.ts ./path/to/file.xlsx ./path/to/file.agent.xlsx
bun run examples/agent-safe-edit-http.ts ./path/to/file.xlsx ./path/to/file.api-agent.xlsx
bun run examples/agent-safe-edit-mcp.ts ./path/to/file.xlsx ./path/to/file.mcp-agent.xlsx
```

Or run the agent workflows from the examples package:

```bash
cd examples
bun run safe-edit ./path/to/file.xlsx ./path/to/file.agent.xlsx
bun run safe-edit:http ./path/to/file.xlsx ./path/to/file.api-agent.xlsx
bun run safe-edit:mcp ./path/to/file.xlsx ./path/to/file.mcp-agent.xlsx
```

`agent-safe-edit.ts` is the golden path for coding agents: trust preflight, inspect, prepare a safe plan, commit with the prepared handle, verify, and print expected JSON fields.
`agent-safe-edit-http.ts` runs the same workflow through the HTTP API fetch surface, and `agent-safe-edit-mcp.ts` runs it through MCP tools; `agent-safe-edit-http.md` and `agent-safe-edit-mcp.md` show API and MCP request shapes, including open-plan preflight, formula assistance, and path-addressed mutations.
`untrusted-workbook-report.md` shows the trust preflight for files from email, downloads, customers, or another agent: CLI `inspect --agent`, API `/trust-report`, and MCP `ascend.trust_report`.

`mcp-setup.md` documents MCP configuration (also summarized in the root README).
