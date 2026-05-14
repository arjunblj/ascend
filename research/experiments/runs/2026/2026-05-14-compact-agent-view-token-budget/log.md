# Compact Agent-View Token Budget

## Question

Can Ascend's compact `agentView` summarize workbook intent under strict, explicit token budgets while preserving shape, column, formula, and truncation facts?

## Hypothesis

Yes. The existing `agentView` already computes the right semantic facts, but it needs a deterministic post-pass that trims evidence-heavy samples and reports the approximate budget outcome.

## External sources checked

- OpenAI prompt caching guide: https://developers.openai.com/api/docs/guides/prompt-caching
- Model Context Protocol resources specification: https://modelcontextprotocol.io/specification/2025-06-18/server/resources
- Notion page content and pagination guide: https://developers.notion.com/guides/data-apis/working-with-page-content
- Notion markdown content API guide: https://developers.notion.com/guides/data-apis/working-with-markdown-content
- Univer spreadsheet MCP guide: https://docs.univer.ai/guides/sheets/features/mcp

## Why this matters to Ascend

Agents need workbook context that is predictable enough to fit inside prompts and MCP resource responses. External agent surfaces point in the same direction: stable context prefixes, paginated or bounded content, and explicit resource retrieval. Ascend should make spreadsheet intent summaries budget-aware instead of forcing every CLI/API/MCP caller to invent its own truncation policy.

## Probe/implementation

- Re-read the existing local probe at `probes/agent-view-budget.ts`, which compared current `agentView` JSON with a projected compact shape against a local private workbook corpus.
- Inspected `WorkbookReadView.agentView`, `AgentViewOptions`, `AgentViewResult`, and existing SDK tests.
- Added `AgentViewOptions.maxApproxTokens`.
- Added optional `AgentViewBudgetInfo` metadata to `AgentViewResult`.
- Implemented a deterministic SDK budget pass that preserves row/column/non-empty/formula facts while trimming sample rows, column sample values, and formula pattern evidence until the approximate JSON-byte token estimate fits where possible.
- Added a targeted SDK test proving budgeted views shrink while retaining shape facts.

## Results

The original private-corpus probe showed a projected compact view could shrink about 5,844 approximate tokens to about 792 while preserving sheet shape, headers, column kinds, formula counts, and representative samples.

The production SDK now exposes the first fold-in of that idea:

```bash
bun test packages/sdk/src/sdk.test.ts -t "agentView"
```

Targeted validation passed with 2 agent-view tests. The new result reports requested approximate tokens, estimated approximate tokens, the estimator, and omitted sample/formula evidence counts.

## Confidence

High for SDK-level value because the change is a deterministic post-processing pass over already computed facts. Medium for end-to-end product value until CLI/API/MCP accept the same option and larger real-workbook fixtures validate budget behavior across tables, pivots, comments, and wide sheets.

## Fold-in decision

Promote to product/DX loop. This cycle folded the core SDK budget contract into production. The next fold-in should expose `maxApproxTokens` through CLI/API/MCP and add docs/examples once the public surface is stable.

## Next question

Can formula span metadata be promoted into a minimal SDK helper for hover/diagnostics/code-action consumers without changing parser semantics?
