# Agent Workflow Ordering Harness

## Question

Can a small deterministic harness prevent Ascend's agent docs from regressing into a workflow that hydrates, inspects, or reads unknown XLSX/XLSM files before `open-plan`?

## Hypothesis

Agent workflow quality should be tested as a tool-call trajectory problem. A cheap docs-surface test can catch the highest-risk ordering regression before it reaches CLI, API, MCP, or bundled `llms` context.

## External sources checked

- Model Context Protocol tools docs: https://modelcontextprotocol.io/docs/concepts/tools
- MCP client best practices: https://modelcontextprotocol.io/docs/develop/clients/client-best-practices
- OpenAI Agents SDK tools docs: https://openai.github.io/openai-agents-python/tools/
- LangChain agent evals docs: https://docs.langchain.com/oss/python/langchain/evals
- LangSmith application-specific evaluation approaches: https://docs.langchain.com/langsmith/evaluation-approaches
- `langchain-ai/agentevals` repository: https://github.com/langchain-ai/agentevals
- TRAJECT-Bench paper: https://arxiv.org/abs/2510.04550

## Why this matters to Ascend

Ascend is trying to be a safer spreadsheet platform for agents, not just a workbook library. If the bundled agent entry points drift and suggest `inspect`, `read`, or `trust-report` before package-level open planning, agents may hydrate active, signed, macro-enabled, or otherwise suspicious packages before choosing a safe open mode. External agent-eval practice treats the ordered sequence of tool calls as a testable artifact, so Ascend should lock down its canonical workflow ordering.

## Probe/implementation

- Inspected `packages/sdk/src/agent-docs.test.ts`, `docs/AGENT_WORKFLOW.md`, `docs/AGENT_API.md`, `llms.txt`, `llms-full.txt`, and the HTTP/MCP safe-edit transcripts.
- Added a deterministic `expectTextOrder` helper to `packages/sdk/src/agent-docs.test.ts`.
- Added a bundled-docs regression test that asserts `open-plan` appears before trust/read/hydration-facing steps in:
  - `docs/AGENT_WORKFLOW.md`
  - `docs/AGENT_API.md`
  - `llms.txt`
  - `llms-full.txt`
  - `examples/agent-safe-edit-http.md`
  - `examples/agent-safe-edit-mcp.md`
- The same test asserts those surfaces mention `reviewBeforeHydration`, so the workflow includes the branch that keeps agents in metadata/trust/package-inventory mode.

## Results

- `bun test packages/sdk/src/agent-docs.test.ts` passed with 7 tests and 99 expectations.
- `bunx biome check packages/sdk/src/agent-docs.test.ts research/experiments/index.md research/experiments/runs/2026/2026-05-14-agent-workflow-ordering-harness/log.md` passed.
- `bunx tsc --build` passed.
- `bun run test:changed` passed with 4960 tests, 1 skip, and 0 failures.

The probe folded into a small production-adjacent regression test rather than another prose-only note.

## Confidence

Medium-high. This does not simulate a real LLM making tool calls, but it locks the bundled machine-readable workflow order that agents and integrations consume. The external evidence supports tool-call trajectory checks, and this is the lowest-cost local guard for the known failure mode.

## Fold-in decision

Folded into the product/DX loop as an SDK docs-surface regression test. No runtime production behavior changed.

## Next question

Can Ascend turn formula oracle routing research into a production-visible mismatch classifier that tells agents which oracle to run first for parser, evaluation, formatting, dependency, or package-preservation disagreements?
