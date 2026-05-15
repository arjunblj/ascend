# Agent View Formula Pattern Provenance

## Question

Can formula-pattern provenance attach compact row/ref examples to agent-view formula summaries so omitted formula patterns can be recovered with narrower reads instead of same-range unbudgeted views?

## Hypothesis

Yes. `agentView` already scans formula cells to normalize formula patterns. It can retain a tiny bounded list of example refs per pattern and include the first omitted pattern's example ref in budget metadata. This gives agents a concrete follow-up cell/range without embedding all formulas.

## External sources checked

- HyperFormula dependency graph docs describe formula cells as addressable graph nodes, reinforcing that formulas need cell provenance, not just aggregate counts: https://hyperformula.handsontable.com/guide/dependency-graph.html
- MCP tool results may link to resources for additional context, supporting progressive disclosure by reference instead of embedding all data: https://modelcontextprotocol.io/docs/concepts/tools
- MCP resources are context objects exposed separately from tool results: https://modelcontextprotocol.io/docs/concepts/resources
- OpenAI Agents SDK context docs emphasize that models only see the conversation/context supplied at call time, making compact follow-up references important: https://openai.github.io/openai-agents-js/guides/context/

## Why this matters to Ascend

The omitted-evidence locator fold-in made sample-row and column-sample recovery concrete, but formula-pattern omissions still only had counts. For agent workflows, "there are omitted formula patterns" is less useful than "the next omitted pattern has an example at D48." Formula provenance moves token-bounded agent view closer to a product-proof recovery flow without claiming full formula-language intelligence.

## Probe/implementation

- Inspected `agentView()` formula pattern aggregation in `packages/sdk/src/read-view.ts`.
- Added bounded `examples` refs to `AgentFormulaPatternInfo`.
- Recorded up to three refs per normalized pattern while scanning formulas.
- Added `nextExampleRef` to omitted formula-pattern budget metadata.
- Updated the recovery proof harness to:
  - read `budget.omittedEvidence.formulaPatterns.nextExampleRef`;
  - run a narrow `agentView` on that cell ref when present;
  - prove the narrow view recovers a formula pattern.
- Added an SDK assertion that formula pattern summaries include example refs.

## Results

Local proof command:

```bash
bun run fixtures/benchmarks/agent-view-recovery-proof.ts
```

Key result:

- `public-formula-stress` has `omittedFormulaPatterns=12`.
- Its budget metadata exposes `nextExampleRef=D48`.
- A narrow follow-up `agentView("Finance", "D48")` recovers a formula pattern.
- All formula-pattern example recoveries in the recovery proof pass.

Boundary:

- Formula examples are provenance hints, not all occurrences.
- Pattern examples are capped at three refs per pattern.
- Formula-pattern recovery is now possible by example ref, but not yet a full "recover every omitted pattern occurrence" planner.

Validation passed:

```bash
bun test packages/sdk/src/sdk.test.ts -t "agentView summarizes a range|agentView applies approximate token budgets"
bun test fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts
bun test apps/api/src/server.test.ts -t "agent-view exposes token budget metadata"
bun test apps/mcp/src/index.test.ts -t "ascend.agent_view exposes token budget metadata"
bunx biome check packages/sdk/src/read-view.ts packages/sdk/src/types.ts packages/sdk/src/sdk.test.ts fixtures/benchmarks/agent-view-recovery-proof.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts fixtures/benchmarks/agent-view-budget-proof.test.ts
bunx tsc --build
bun run test:changed
```

## Confidence

High that formula-pattern examples are a low-cost, useful provenance primitive. Medium that the exact cap and shape are final; product may later want a separate recovery planner that groups omitted pattern examples by row/range.

## Fold-in decision

Fold into production as a compact agent-view provenance enhancement. Keep the claim bounded to example-driven recovery, not complete formula occurrence recovery.

## Next question

Can token-bounded agent view now graduate into the release proof digest index as a third artifact, or should it wait for a public product example that demonstrates sample-row, column-sample, and formula-pattern recovery end to end?
