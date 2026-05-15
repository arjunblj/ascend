# Write Error Structured Journal Proof

## Question

Do API and MCP write errors preserve structured journal build-failure evidence for agents, or do they collapse failed applies down to the first validation error?

## Hypothesis

They should preserve the structured apply result in error details. Agents need the journal's `JOURNAL_BUILD_FAILED` issue, undo policy, and surface/reason metadata to repair or explain failed writes.

## External sources checked

- JSON-RPC 2.0 includes an error `data` member for structured additional error information: https://www.jsonrpc.org/specification
- Model Context Protocol tool results use `isError: true` for tool execution errors and support structured result content: https://modelcontextprotocol.io/docs/concepts/tools
- Language Server Protocol diagnostics preserve structured codes and related diagnostic data for tooling: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/

## Why this matters to Ascend

The auditable mutation claim is not limited to successful commits. If an agent submits an invalid operation with `journal: true`, the API/MCP surfaces should return enough structured evidence to explain that rollback evidence could not be built.

## Probe/implementation

Finished the in-flight API/MCP fold-in:

- API `/write` now attaches the full `apply` result to validation-error details when apply fails.
- MCP `ascend.write` now attaches the full `apply` result to structured tool-error details when apply fails.
- Added API and MCP tests using invalid `clearRange` input `A1:` with `journal: true`.

Commands run:

```bash
bun test apps/api/src/server.test.ts -t "write errors preserve structured journal build failures"
bun test apps/mcp/src/index.test.ts -t "ascend.write errors preserve structured journal build failures"
bunx biome check apps/api/src/server.ts apps/api/src/server.test.ts apps/mcp/src/index.ts apps/mcp/src/index.test.ts
bunx tsc --build
```

## Results

- API focused test passed: 1 test, 3 assertions.
- MCP focused test passed: 1 test, 3 assertions.
- Biome passed for the touched API/MCP files.
- `bunx tsc --build` passed.
- Both surfaces now expose `apply.journal` with:
  - `supported=false`
  - `exact=false`
  - issue code `JOURNAL_BUILD_FAILED`
  - surface `package-parts`
  - reason `journal-build-failed`
  - undo policy reason `build-failed`

## Confidence

High for failed apply/journal build failures on API and MCP writes. Medium for all write-error classes until the same shape is tested for parse errors, partial workbook failures, and recalc failures.

## Fold-in decision

Promote to product/correctness loop. This is a narrow API/MCP structured-error fix for existing write surfaces, not a new mutation surface.

## Next question

Should CLI write errors preserve the same `apply.journal` details, or does CLI already print enough structured JSON for failed applies?
