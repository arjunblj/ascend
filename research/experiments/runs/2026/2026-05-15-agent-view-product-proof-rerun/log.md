# Agent View Product Proof Rerun

## Question

Can the token-bounded agent view claim move from "needs recovery proof" to "product-proof backed" using existing SDK/CLI/API/MCP surfaces and tracked budget/recovery harnesses?

## Hypothesis

Yes. The existing agent-view budget and omitted-evidence recovery harnesses should show deterministic summaries, counted omissions, compact recovery locators, and representative formula-pattern recovery without adding another product surface.

## External sources checked

- Univer MCP getting started: https://docs.univer.ai/guides/sheets/getting-started/mcp
- Univer MCP feature guide: https://docs.univer.ai/guides/sheets/features/mcp
- Notion MCP guide: https://developers.notion.com/guides/mcp/mcp
- Notion MCP supported tools: https://developers.notion.com/guides/mcp/mcp-supported-tools
- OpenAI Agents SDK context management: https://openai.github.io/openai-agents-js/guides/context/
- MCP tools concept: https://modelcontextprotocol.wiki/en/docs/concepts/tools
- MCP resources concept: https://modelcontextprotocol.wiki/en/docs/concepts/resources

## Why this matters to Ascend

Spreadsheet agents need bounded context that is deterministic, recoverable, and honest about omissions. Competitors increasingly expose agent spreadsheet operations through MCP, but Ascend can differentiate by making local workbook summaries explicit about budgets, omitted evidence, and recovery paths.

## Probe/implementation

- Ran `git status --short --branch`; no tracked production diff was present.
- Inspected tracked agent-view budget and recovery harnesses.
- Ran local probes:

```bash
bun run fixtures/benchmarks/agent-view-budget-proof.ts
bun run fixtures/benchmarks/agent-view-recovery-proof.ts
```

- Validated harnesses and surfaces:

```bash
bun test fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts
bun test packages/sdk/src/sdk.test.ts -t "agentView applies approximate token budgets"
bun test apps/cli/src/cli.test.ts -t "agent-view --tokens"
bun test apps/api/src/server.test.ts -t "agent-view exposes token budget metadata"
bun test apps/mcp/src/index.test.ts -t "agent_view exposes token budget metadata"
```

- Updated:
  - `research/experiments/syntheses/2026-05-agent-context-contracts.md`
  - `research/experiments/syntheses/2026-05-release-claim-board.md`

## Results

Budget proof covered 5 workbook shapes:

- dense table;
- wide sparse sheet;
- formula-heavy generated sheet;
- metadata-heavy generated sheet;
- public formula-stress workbook.

Results:

- all cases were deterministic;
- all preserved sheet shape;
- all truncations had omission counters;
- full-to-budgeted compression ratios ranged from `0.206` to `0.983`;
- two cases exceeded the requested budget: wide sparse and public formula-stress.

Recovery proof:

- same-range unbudgeted recovery was exact for every case;
- omitted sample-row and column-sample locators were present and exact;
- narrower sample-row recovery was exact for every case;
- public formula-stress included a formula-pattern example ref, `D48`, and recovered representative formula evidence.

Cross-surface validation passed for SDK, CLI, API, and MCP token-budget metadata.

## Confidence

High for product-proof backed wording around deterministic summaries, omission counters, compact locators, and same-range/narrow-read recovery. Medium for release headline wording because the proof still needs a short product example that shows an agent using the omitted-evidence locator to perform the follow-up read.

## Fold-in decision

Promote to product/DX proof packaging, not production. No new surfaces. Keep exact-token and complete-evidence claims out of release copy.

## Next question

Can the retained viewport patch history claim be narrowed into a product example that shows SDK patch success, API/MCP `changedSince` recovery, and explicit CLI exclusion without adding CLI `changedSince`?
