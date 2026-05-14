# Ascend examples

Run from the **repository root** with Bun:

```bash
bun run examples/create-from-scratch.ts
bun run examples/read-modify-save.ts ./path/to/file.xlsx
bun run examples/formula-eval.ts
bun run examples/csv-convert.ts
bun run examples/batch-ops.ts
bun run examples/agent-safe-edit.ts ./path/to/file.xlsx ./path/to/file.agent.xlsx
```

`agent-safe-edit.ts` is the golden path for coding agents: inspect, prepare a safe plan, commit with the prepared handle, verify, and print expected JSON fields.

`mcp-setup.md` documents MCP configuration (also summarized in the root README).
