# Agent View Omitted Locator Metadata

## Question

Can compact omitted-evidence locator metadata be added to budgeted agent views without bloating the response beyond the token savings it enables?

## Hypothesis

Yes, with a narrow scope. Budgeted agent views can carry bounded omitted sample-row and column-sample locators inside the existing `budget` object. The locators are small enough for most cases and let clients perform narrower follow-up reads for omitted sample rows. Formula-pattern recovery still needs richer provenance before it can be fully automated.

## External sources checked

- MCP tools can return resource links for additional context, reinforcing progressive disclosure instead of one giant tool result: https://modelcontextprotocol.io/docs/concepts/tools
- OpenAI Agents SDK context docs emphasize that the model only sees conversation-history data at call time, making explicit context selection important: https://openai.github.io/openai-agents-js/guides/context/
- OpenAI Agents SDK usage docs track context-window consumption through request usage: https://openai.github.io/openai-agents-python/usage/
- Notion MCP docs position MCP as a way for AI tools to interact with workspace data, making returned context shape part of the product contract: https://developers.notion.com/guides/mcp/overview

## Why this matters to Ascend

The previous recovery proof found that same-range unbudgeted recovery was exact but budget metadata was count-only. That left agents without enough information to choose a smaller follow-up read. Compact locators turn omission counters into a usable recovery path while keeping the release claim honest about structural floors and formula-pattern limits.

## Probe/implementation

- Ran `git status --short --branch` before production edits.
- Inspected `budgetAgentViewResult()` and `AgentViewBudgetInfo`.
- Added `AgentViewOmittedEvidenceInfo` and nested locator types to `packages/sdk/src/types.ts`.
- Updated `packages/sdk/src/read-view.ts` so budgeted agent views can include:
  - omitted sample-row count plus first/last omitted row;
  - up to two omitted column-sample hints plus remaining-column count;
  - omitted formula-pattern count and first omitted index, without embedding pattern text.
- Kept locator metadata bounded so it does not become a second full evidence payload.
- Updated SDK/API/MCP tests that already exercise agent-view budget metadata.
- Updated the agent-view recovery proof harness to construct a narrower follow-up range from sample-row locators and prove omitted sample rows are recoverable.

## Results

Local proof command:

```bash
bun run fixtures/benchmarks/agent-view-recovery-proof.ts
```

Results:

- All five cases expose omitted-evidence locators.
- Same-range unbudgeted recovery remains exact.
- Sample-row locators are exact for all five cases.
- Column-sample locator counts are exact for all five cases, with bounded visible column hints.
- Narrow sample-row follow-up reads recovered the omitted sample rows in all five cases.
- Wide sparse remains a structural-floor boundary: the locator range can still be the full requested range when omitted sample rows are far apart.
- Public formula stress still slightly exceeds the 640-token request after locator metadata. This is acceptable for the current claim because budget metadata is approximate and structural floors are explicit; it should not be advertised as strict budget adherence.

Validation passed:

```bash
bun test packages/sdk/src/sdk.test.ts -t "agentView applies approximate token budgets"
bun test apps/api/src/server.test.ts -t "agent-view exposes token budget metadata"
bun test apps/mcp/src/index.test.ts -t "ascend.agent_view exposes token budget metadata"
bun test fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts
bunx biome check packages/sdk/src/read-view.ts packages/sdk/src/types.ts packages/sdk/src/sdk.test.ts apps/api/src/server.test.ts apps/mcp/src/index.test.ts fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts
```

## Confidence

High that the locator metadata is a useful product/DX fold-in. Medium that the exact locator schema is final: formula patterns still have no source locations, and agents may need policy for choosing between sample-row range reads, column-focused reads, or same-range unbudgeted recovery.

## Fold-in decision

Fold into production as a small SDK/API/MCP-compatible budget metadata extension. Keep token-bounded agent view below the top two release claims until a product proof documents recovery policy and public examples.

## Next question

Can formula-pattern provenance attach compact row/ref examples to agent-view formula summaries so omitted formula patterns can be recovered with narrower reads instead of same-range unbudgeted views?
