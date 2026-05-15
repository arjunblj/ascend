# Agent View Product Example Gate

## Question

Should token-bounded agent view remain below the top two release claims until omitted-evidence recovery is product-defined, or is the current proof strong enough for a product example without release-index promotion?

## Hypothesis

The current proof is strong enough for a product example because locator metadata and formula-pattern example refs now provide concrete recovery paths. It should still stay out of the release proof digest index because that index is scoped to the top two release claims.

## External sources checked

- Model Context Protocol specification: https://modelcontextprotocol.io/specification/2024-11-05/index
- Univer MCP guide: https://docs.univer.ai/guides/sheets/getting-started/mcp
- Microsoft Graph Excel workbook APIs: https://learn.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0

## Why this matters to Ascend

Agents need bounded spreadsheet context that is honest about missing evidence. The useful product claim is not "we fit every workbook into a prompt"; it is "we return deterministic omission metadata and give the agent an explicit recovery path."

## Probe/implementation

- Inspected current agent-context synthesis and tracked proof harnesses:
  - `research/experiments/syntheses/2026-05-agent-context-contracts.md`
  - `fixtures/benchmarks/agent-view-budget-proof.ts`
  - `fixtures/benchmarks/agent-view-recovery-proof.ts`
- Reran:
  - `bun run fixtures/benchmarks/agent-view-budget-proof.ts`
  - `bun run fixtures/benchmarks/agent-view-recovery-proof.ts`
  - `bun test fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts`
- Reran surface validation:
  - `bun test packages/sdk/src/sdk.test.ts -t "agentView applies approximate token budgets"`
  - `bun test apps/cli/src/cli.test.ts -t "agent-view --tokens"`
  - `bun test apps/api/src/server.test.ts -t "agent-view exposes token budget metadata"`
  - `bun test apps/mcp/src/index.test.ts -t "agent_view exposes token budget metadata"`
- Updated the agent-context synthesis with a product-example gate and release-index boundary.

## Results

Budget proof stayed deterministic across five cases:

| Case | Requested | Full tokens | Budgeted tokens | Within budget | Omitted evidence | Shape preserved |
| --- | ---: | ---: | ---: | --- | ---: | --- |
| dense-table | 512 | 2010 | 415 | true | 24 | true |
| wide-sparse | 384 | 866 | 851 | false | 7 | true |
| formula-heavy | 512 | 1125 | 494 | true | 11 | true |
| metadata-heavy | 448 | 1560 | 346 | true | 20 | true |
| public-formula-stress | 640 | 1705 | 658 | false | 63 | true |

Recovery proof stayed concrete:

- same-range unbudgeted recovery exact for all cases;
- compact omitted-evidence locators present for all cases;
- sample-row locators exact for all cases;
- column-sample locators exact for all cases;
- narrow sample-row recovery exact for all cases;
- formula-pattern example recovery exact for the public formula-stress case.

## Confidence

High that token-bounded agent view is product-example ready. Medium that it should become a release-index artifact; the index should stay focused on safe-open and package-action until product defines persisted context-contract artifact policy.

## Fold-in decision

Promote to topic synthesis and product example. Do not add a new SDK/CLI/API/MCP surface. Do not add agent-view to `fixtures/benchmarks/release-proof-index.ts` yet.

## Next question

Can retained viewport patch history receive the same product-example gate, with SDK patch success, API/MCP compact `changedSince` recovery, and explicit CLI exclusion?
