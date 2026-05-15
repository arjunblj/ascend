# Agent View Index Promotion Gate

## Question

Can the token-bounded agent view claim be tied to the release proof digest index now, or should it remain below the top two until product defines recovery flows for omitted evidence?

## Hypothesis

Hold it out of the release proof index for now. The agent-view harness proves deterministic budget metadata and omission counters, but the release index is currently scoped to the top two claims. Token-bounded agent view still needs a product recovery story for omitted evidence before it should sit beside safe-open and package-action proof artifacts.

## External sources checked

- MCP sampling schema includes `maxTokens` as an explicit sampling limit and `maxTokens` as a stop reason: https://modelcontextprotocol.io/specification/2025-06-18/schema
- MCP sampling draft frames tool/model sampling as cross-provider context management: https://modelcontextprotocol.io/specification/draft/client/sampling
- Notion MCP docs position MCP as an agent workflow over workspace data, reinforcing that tool output shape is product-facing context: https://developers.notion.com/guides/mcp/mcp
- OpenAI Agents SDK usage docs track request usage for cost and context-window consumption: https://openai.github.io/openai-agents-python/usage/

## Why this matters to Ascend

Token-bounded agent view is credible, but promoting it into the release proof index would imply it is one of the release's top proof artifacts. The claim board intentionally ranked safe unknown workbook opening and auditable package-part mutation above it. Ascend should avoid broadening the release proof set until the product loop defines how users recover omitted rows, omitted column samples, and omitted formula patterns.

## Probe/implementation

- Ran `git status --short --branch`; only untracked research/private/tmp files were present.
- Inspected:
  - `fixtures/benchmarks/agent-view-budget-proof.ts`
  - `fixtures/benchmarks/agent-view-budget-proof.test.ts`
  - `fixtures/benchmarks/release-proof-index.ts`
  - `research/experiments/syntheses/2026-05-release-claim-board.md`
  - current SDK/CLI/API/MCP agent-view budget tests by search.
- Reran the current agent-view proof:

```bash
bun run fixtures/benchmarks/agent-view-budget-proof.ts
```

- Kept the release proof index scoped to `safe-open-proof` and `package-action-proof`.
- Updated the release claim board to make the promotion gate explicit.

## Results

The harness still supports the conservative claim:

- 5 mixed workbook shapes ran: dense table, wide sparse, formula-heavy, metadata-heavy, public formula stress.
- All outputs were deterministic.
- Shape facts were preserved in all cases.
- Omission counters were present for all truncated outputs.
- Unbudgeted token estimates matched the full agent-view estimate.

The hold decision is based on two boundaries:

- `wide-sparse` requested 384 approximate tokens but emitted 791 because preserving 26 column summaries is the structural floor.
- The current recovery instruction is honest but not product-packaged: omitted evidence must be recovered through narrower reads or unbudgeted views. The release index should not promote the claim until that recovery path is surfaced as a proof flow.

## Confidence

High that token-bounded agent view should remain out of the top-claim release proof index for this block. Medium that the next product proof should be "recovery from omitted evidence" rather than another budget knob, because current surfaces already expose the core budget metadata.

## Fold-in decision

Promote to topic synthesis only. Do not add agent-view to `fixtures/benchmarks/release-proof-index.ts` yet. Keep the release index limited to safe-open and package-action proof artifacts.

## Next question

Can a product-proof harness demonstrate omitted-evidence recovery by taking an agent-view budget result, selecting a narrower range from omission metadata, and proving the recovered view contains the omitted formula/sample evidence?
