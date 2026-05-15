# Agent view token budget surfaces

## Question

Can the token-bounded agent-view claim be made product-real across CLI, API, and MCP without changing the SDK summarization model?

## Hypothesis

Yes. The SDK already supports `maxApproxTokens` and returns budget metadata with omission counters. The missing fold-in is to expose that option through the existing CLI/API/MCP `agent-view` surfaces.

## External sources checked

- [Univer MCP spreadsheet docs](https://docs.univer.ai/guides/sheets/getting-started/mcp) show an agent-native spreadsheet workflow where tools are automatically available to LLM clients.
- [MCP resources documentation](https://modelcontextprotocol.io/docs/concepts/resources) frames resources as model context and notes pagination support, which supports bounded spreadsheet context as a product contract.
- [MCP pagination specification](https://modelcontextprotocol.io/specification/draft/server/utilities/pagination) reinforces cursor/page-sized context as a first-class agent protocol pattern.
- [Notion pagination docs](https://developers.notion.com/reference/pagination) are a mature API contrast for explicit page-size/cursor bounded data access.
- [OpenAI prompt caching docs](https://platform.openai.com/docs/guides/prompt-caching) reinforce why stable, bounded prompt inputs matter for latency and cost management.

## Why this matters to Ascend

The claim ladder says "token-bounded agent view" is credible today but should be packaged as a product claim. Agents need the same budget knob whether they use CLI, HTTP, or MCP. Without surfacing `maxApproxTokens`, only SDK users can request deterministic truncation metadata.

## Probe/implementation

- Inspected `packages/sdk/src/read-view.ts` and confirmed budgeted agent views already preserve row/column/non-empty/formula facts while trimming sample rows, column sample values, and formula patterns.
- Added `--tokens <count>` to `ascend agent-view`.
- Added `maxApproxTokens` to `POST /agent-view`.
- Added `maxApproxTokens` to MCP `ascend.agent_view`.
- Updated the API/MCP contract docs where the public surface changed.
- Added focused CLI, API, and MCP tests that assert budget metadata and omission counters.
- Ran a generated-workbook local probe comparing full and budgeted agent views.

## Results

Focused validation passed:

```bash
bun test apps/cli/src/cli.test.ts -t "agent-view --tokens"
bun test apps/api/src/server.test.ts -t "agent-view exposes token budget"
bun test apps/mcp/src/index.test.ts -t "agent_view exposes token budget"
bun test packages/sdk/src/sdk.test.ts -t "agentView applies approximate token budgets"
```

Generated-workbook probe:

| Full approximate tokens | Requested approximate tokens | Budgeted approximate tokens | Omitted sample rows | Omitted column sample values | Shape retained |
| ---: | ---: | ---: | ---: | ---: | --- |
| 3050 | 640 | 587 | 11 | 32 | 40 rows, 8 columns, 320 non-empty cells |

## Confidence

High that token-bounded agent view now reaches the main product surfaces. Medium for public claim strength until docs/examples show recovery guidance for omitted evidence and benchmark the behavior on real table/formula/metadata-heavy fixtures.

## Fold-in decision

Promote to product/DX loop. This cycle folded the budget knob into CLI/API/MCP and proved it with focused tests. The next loop should package examples and boundaries, not add more knobs.

## Next question

Can retained viewport patch history be packaged as an honest product claim with recovery guidance for invalid, stale, expired, and mismatched tokens?
