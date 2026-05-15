# Agent View Omitted Evidence Recovery

## Question

Can a product-proof harness demonstrate omitted-evidence recovery by taking an agent-view budget result, selecting a narrower range from omission metadata, and proving the recovered view contains the omitted formula/sample evidence?

## Hypothesis

Partially. Ascend can prove exact same-range recovery by rerunning `agentView` without a budget, but it cannot yet prove automatic narrower-range recovery from the budget response alone because budget metadata carries omission counts, not omitted row/column/formula locations.

## External sources checked

- MCP schema includes `maxTokens` and `maxTokens` stop reasons, which makes bounded tool/model context explicit: https://modelcontextprotocol.io/specification/2025-06-18/schema
- MCP sampling draft frames context and token limits as cross-provider sampling concerns: https://modelcontextprotocol.io/specification/draft/client/sampling
- Notion MCP docs position MCP as a way for assistants to interact with workspace data, making returned context shape a product contract: https://developers.notion.com/guides/mcp/overview
- OpenAI Agents SDK usage docs track request usage for cost and context-window consumption: https://openai.github.io/openai-agents-python/usage/

## Why this matters to Ascend

The release claim board deliberately keeps token-bounded agent view behind the top two claims. The missing proof is not another budget knob; it is recovery from omitted evidence. Agents need to know whether a budgeted summary can lead them to the missing rows, column samples, or formula patterns without falling back to full context every time.

## Probe/implementation

- Inspected current budget implementation in `packages/sdk/src/read-view.ts`.
- Confirmed `AgentViewBudgetInfo` exposes:
  - requested, estimated, and unbudgeted approximate tokens;
  - truncation flag;
  - omitted sample-row, column-sample-value, and formula-pattern counts.
- Confirmed it does not expose omitted row/column/formula locations.
- Added `fixtures/benchmarks/agent-view-recovery-proof.ts`.
  - Reuses the existing mixed-shape agent-view budget proof cases.
  - Compares budgeted view shape against same-range unbudgeted recovery.
  - Records whether budget metadata has omitted-location fields.
  - Emits Markdown/JSON proof output.
- Added `fixtures/benchmarks/agent-view-recovery-proof.test.ts`.

## Results

Local proof command:

```bash
bun run fixtures/benchmarks/agent-view-recovery-proof.ts
```

| Case | Range | Within budget | Omitted rows | Omitted values | Omitted formulas | Has locations | Same-range recovered |
| --- | --- | --- | ---: | ---: | ---: | --- | --- |
| dense-table | Sheet1!A1:H40 | true | 8 | 16 | 0 | false | true |
| wide-sparse | Sheet1!A1:Z40 | false | 3 | 4 | 0 | false | true |
| formula-heavy | Sheet1!A1:F60 | true | 7 | 4 | 0 | false | true |
| metadata-heavy | Sheet1!A1:F24 | true | 8 | 12 | 0 | false | true |
| public-formula-stress | Finance!A1:L62 | true | 8 | 43 | 6 | false | true |

Findings:

- Same-range unbudgeted recovery is exact for all five cases.
- Budget metadata is count-only for all five cases.
- The wide-sparse structural-floor boundary remains: 384 requested approximate tokens emitted 791 because shape-preserving column summaries dominate.
- Automatic narrower recovery is not currently product-proofable because the budget response does not identify omitted rows, columns, or formula patterns.

## Confidence

High for the current boundary: unbudgeted recovery works, automated narrower recovery needs locator metadata. Medium for the future design: omitted-location metadata may increase response size, so the product loop should prototype compact locator hints before adding them to production surfaces.

## Fold-in decision

Fold into benchmark/proof tooling only. Do not add agent-view to the release proof digest index yet. Promote the next product/DX fold-in candidate as compact omitted-evidence locator metadata, with tests proving it can drive narrower follow-up reads.

## Next question

Can compact omitted-evidence locator metadata be added to budgeted agent views without bloating the response beyond the token savings it enables?
