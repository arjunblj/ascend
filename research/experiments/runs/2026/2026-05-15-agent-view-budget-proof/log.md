# Agent View Budget Proof

## Question

Can the agent-view claim be tightened with fixture-backed token budget proof that compares compact agent views against raw sheet JSON for the same workbook?

## Hypothesis

Yes. Ascend already preserves range shape facts in `agentView`; adding unbudgeted token cost metadata to budgeted views gives agents and product docs an honest compression proof without embedding raw sheet payloads.

## External sources checked

- OpenAI prompt caching docs recommend structuring repeated/static context separately from variable context and expose token/cost implications for long prompts. https://platform.openai.com/docs/guides/prompt-caching/prompt-caching
- OpenAI latency optimization docs call out context filtering/pruning as a latency lever. https://platform.openai.com/docs/guides/latency-optimization
- MCP resource docs explicitly allow paginated resources, matching bounded spreadsheet inspection rather than one giant payload. https://modelcontextprotocol.io/docs/concepts/resources
- MCP schema docs include `maxTokens` on sampling messages, reinforcing token-budgeted tool responses. https://modelcontextprotocol.io/specification/2025-06-18/schema
- Notion API property-item docs paginate large property values and relations instead of forcing full inline records. https://developers.notion.com/reference/page-property-values

## Why this matters to Ascend

"Token-bounded agent view" is a product-shaped claim. It means an agent can inspect workbook intent under a budget while retaining shape facts like row count, column count, non-empty count, formula count, column summaries, and truncation evidence. The proof should quantify what was saved, not just say "truncated."

## Probe/implementation

Local probe created a 30x6 synthetic workbook in memory and compared:

- raw `streamRange('Sheet1', 'A1:F30')` JSON
- full `agentView('Sheet1', 'A1:F30')`
- budgeted `agentView(..., { maxApproxTokens })`

Probe result:

```json
{
  "rawApproxTokens": 3983,
  "fullAgentViewApproxTokens": 1612,
  "target": 886,
  "budget": {
    "requestedApproxTokens": 886,
    "estimatedApproxTokens": 833,
    "unbudgetedApproxTokens": 1612,
    "estimator": "json-bytes-div-4",
    "truncated": true,
    "omittedSampleRows": 5,
    "omittedColumnSampleValues": 0,
    "omittedFormulaPatterns": 0
  },
  "shape": {
    "rows": 30,
    "cols": 6,
    "nonEmpty": 180
  }
}
```

Fold-in:

- Added `unbudgetedApproxTokens` to `AgentViewBudgetInfo`.
- `budgetAgentViewResult()` now records the full unbudgeted agent-view estimate before trimming samples/patterns.
- Updated SDK budget tests to assert the recorded unbudgeted estimate equals the full view estimate.

## Results

The budget proof now says:

- raw sheet JSON estimate: 3983 tokens
- full agent-view estimate: 1612 tokens
- budgeted agent-view estimate: 833 tokens
- shape facts retained: 30 rows, 6 columns, 180 non-empty cells
- truncation evidence: 5 sample rows omitted

Validation passed:

- `bun test packages/sdk/src/sdk.test.ts -t "agentView applies approximate token budgets"`
- `bun test apps/cli/src/cli.test.ts -t "agent-view --tokens returns budget metadata"`
- `bun test apps/api/src/server.test.ts -t "agent-view exposes token budget metadata"`
- `bun test apps/mcp/src/index.test.ts -t "ascend.agent_view exposes token budget metadata"`
- `bunx biome check packages/sdk/src/read-view.ts packages/sdk/src/types.ts packages/sdk/src/sdk.test.ts`
- `bunx tsc --build`

Validation blocked by unrelated dirty journal worktree changes:

- `bun run test:changed` fails in journal/API/MCP/agent workflow tests because dirty journal changes add `surface` and `reason` fields to issue objects while tests still expect the older exact shape.
- This cycle touched only `packages/sdk/src/read-view.ts`, `packages/sdk/src/types.ts`, and `packages/sdk/src/sdk.test.ts`.

## Confidence

High for the budget metadata and proof shape. Medium for "token" exactness because the estimator remains documented as `json-bytes-div-4`, not model-specific tokenization.

## Fold-in decision

Promote to product/DX loop. This is now a small production fold-in that makes the token-bounded agent-view claim easier to prove.

## Next question

Can columnar scan sidecars show a performance proof on real table/range workloads without changing workbook semantics?
