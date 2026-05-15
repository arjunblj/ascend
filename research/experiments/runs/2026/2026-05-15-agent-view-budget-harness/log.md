# Agent View Budget Harness

## Question

Can a tracked proof harness show deterministic budget adherence and omitted-evidence recovery for token-bounded agent views across mixed workbook shapes without changing agent-view surfaces?

## Hypothesis

Yes, with one honest boundary. Existing `agentView(..., { maxApproxTokens })` already returns requested budgets, estimated budgets, unbudgeted estimates, truncation flags, and omission counters. A harness can prove determinism, shape preservation, compression, and omission accounting across dense, sparse, formula-heavy, metadata-heavy, and public formula workbooks. It should also record cases where structural summary facts exceed the requested budget.

## External sources checked

- MCP sampling docs include maximum-token controls for model calls, reinforcing explicit token budgeting at agent boundaries: https://modelcontextprotocol.io/specification/draft/client/sampling
- Anthropic introduced MCP as a standard for connecting assistants to external data/tools, making bounded context retrieval a core product concern: https://www.anthropic.com/news/model-context-protocol
- Notion MCP docs position MCP as an agent workflow surface over workspace data: https://developers.notion.com/guides/mcp/mcp
- OpenAI Agents SDK context docs expose run usage and context management as first-class agent concerns: https://openai.github.io/openai-agents-js/guides/context/

## Why this matters to Ascend

The release claim board says token-bounded agent view is credible but needs product proof packaging. Agents need workbook context that is deterministic, bounded, and recoverable. The product claim should not be "we always fit any workbook into any budget"; it should be "we preserve shape facts, expose approximate budget metadata, count omitted evidence, and make the structural floor visible."

## Probe/implementation

- Inspected `packages/sdk/src/read-view.ts`, SDK/CLI/API/MCP agent-view tests, agent context contracts, and prior budget proof logs.
- Added `fixtures/benchmarks/agent-view-budget-proof.ts`.
  - Generates dense, wide sparse, formula-heavy, and metadata-heavy workbooks.
  - Opens public `fixtures/xlsx/poi/formula_stress_test.xlsx`.
  - Compares full and budgeted agent views.
  - Asserts deterministic repeated output, shape preservation, unbudgeted estimate recording, compression ratio, and omission counters.
  - Emits Markdown by default and JSON with `--json`.
- Added `fixtures/benchmarks/agent-view-budget-proof.test.ts`.
  - Asserts mixed workbook coverage.
  - Asserts determinism, shape preservation, omission counters, compression, and the explicit structural-floor boundary.
  - Asserts claim-safe report wording.

## Results

Local proof command:

```bash
bun run fixtures/benchmarks/agent-view-budget-proof.ts
```

| Case | Fixture | Range | Requested | Full tokens | Budgeted tokens | Unbudgeted recorded | Ratio | Within budget | Deterministic | Truncated | Omitted rows | Omitted values | Omitted formulas | Shape preserved | Metadata signals |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- | --- | --- | ---: | ---: | ---: | --- | --- |
| dense-table | generated dense table | Sheet1!A1:H40 | 512 | 2010 | 355 | 2010 | 0.177 | true | true | true | 8 | 16 | 0 | true | omittedEvidence=24 |
| wide-sparse | generated wide sparse sheet | Sheet1!A1:Z40 | 384 | 866 | 791 | 866 | 0.913 | false | true | true | 3 | 4 | 0 | true | omittedEvidence=7 |
| formula-heavy | generated formula sheet | Sheet1!A1:F60 | 512 | 1111 | 421 | 1111 | 0.379 | true | true | true | 7 | 4 | 0 | true | omittedEvidence=11 |
| metadata-heavy | generated metadata sheet | Sheet1!A1:F24 | 448 | 1560 | 286 | 1560 | 0.183 | true | true | true | 8 | 12 | 0 | true | comments=1, validations=1, conditionalFormats=1, omittedEvidence=20 |
| public-formula-stress | `fixtures/xlsx/poi/formula_stress_test.xlsx` | Finance!A1:L62 | 640 | 1648 | 639 | 1648 | 0.388 | true | true | true | 8 | 43 | 6 | true | omittedEvidence=57 |

Dead-end/boundary found:

- Wide sparse sheets can have a structural summary floor above a very small budget. The budgeted view still compressed from 866 to 791 approximate tokens and counted omissions, but did not fit the requested 384-token budget because preserving 26 column summaries dominates. The claim must say budget metadata and omission counters are deterministic, not that every requested budget is achievable.
- Agent view is not a replacement for metadata inspection. The metadata-heavy case records comments/validations/conditional formats as inspect signals, but `agentView` remains a cell/range intent summary.

Validation passed:

- `bun test fixtures/benchmarks/agent-view-budget-proof.test.ts`
- `bunx biome check fixtures/benchmarks/agent-view-budget-proof.ts fixtures/benchmarks/agent-view-budget-proof.test.ts`
- `bunx tsc --build`
- `bun run test:changed` (5004 pass, 1 skip)

## Confidence

High that the harness proves deterministic budget behavior and honest structural-floor boundaries. Medium for product proof breadth because public fixture coverage is currently formula-heavy only; product may still want public wide/table/metadata workbook fixtures before external release copy.

## Fold-in decision

Fold into the product/DX proof loop as a tracked benchmark/proof harness. Do not add new agent-view options or surfaces. Update claim wording to mention approximate estimates, structural floors, and omitted-evidence recovery.

## Next question

Can retained viewport patch history get the same proof treatment: deterministic fixture cases for valid token, stale token, expired token, projection change, and bounded retention miss without adding a collaboration claim?
