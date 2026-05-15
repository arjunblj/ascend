# 2026-05 Agent Context Contracts

Date: 2026-05-14

## External References

- [Model Context Protocol specification](https://modelcontextprotocol.io/specification/2024-11-05/index) frames tools, resources, and prompts as protocol primitives for connecting LLM applications to external data. Ascend's agent view should stay a bounded evidence result, not an unbounded workbook dump.
- [Univer MCP guide](https://docs.univer.ai/guides/sheets/getting-started/mcp) shows spreadsheet operations exposed to agents through MCP. Ascend's differentiation is local, deterministic budget metadata and recovery locators.
- [Microsoft Graph Excel workbook APIs](https://learn.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0) expose range-level workbook access. Ascend should make its range summaries recoverable through narrower follow-up reads when evidence is omitted.

## Product Claims

| Claim | Status | Proof now | Boundary |
| --- | --- | --- | --- |
| Token-bounded agent view | Product-proof backed | SDK/CLI/API/MCP accept token budgets and return estimated tokens, omission counters, compact omitted-evidence locators, and formula-pattern example refs | Estimates are approximate; omitted evidence is absent by design and must be recovered through narrower reads or unbudgeted views |
| Retained viewport patch history | Product-proof backed with CLI excluded | SDK interactive viewports return retained patches or explicit invalidation reasons; API/MCP compact reads return `changedSince` recovery windows and invalidation metadata | Bounded per-window/per-viewport history, not CRDT collaboration, multi-writer sync, or unbounded event sourcing |

## Required Examples

| Example | Surface | Must show |
| --- | --- | --- |
| Budgeted sheet summary | CLI/API/MCP | requested tokens, estimated tokens, truncated flag, omitted sample rows, omitted column values |
| Stable compact poll | SDK/API/MCP read | first full compact window, unchanged poll, one-cell patch |
| Recovery from invalid token | SDK/API/MCP read | `base-token-invalid` and use returned full window |
| Recovery from expired token | SDK compact and SDK interactive | `base-token-expired` and use returned snapshot/window |
| Recovery from metadata/layout change | SDK interactive | `viewport-invalidated` and use returned snapshot |

## Token-Bounded Agent View Proof Rerun

Latest rerun: 2026-05-15T03:43:00Z.

Commands:

```bash
bun run fixtures/benchmarks/agent-view-budget-proof.ts
bun run fixtures/benchmarks/agent-view-recovery-proof.ts
bun test fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts
bun test packages/sdk/src/sdk.test.ts -t "agentView applies approximate token budgets"
bun test apps/cli/src/cli.test.ts -t "agent-view --tokens"
bun test apps/api/src/server.test.ts -t "agent-view exposes token budget metadata"
bun test apps/mcp/src/index.test.ts -t "agent_view exposes token budget metadata"
```

Budget proof:

| Case | Range | Requested | Full tokens | Budgeted tokens | Ratio | Within budget | Omitted rows | Omitted values | Omitted formulas | Shape preserved |
| --- | --- | ---: | ---: | ---: | ---: | --- | ---: | ---: | ---: | --- |
| dense-table | Sheet1!A1:H40 | 512 | 2010 | 415 | 0.206 | true | 8 | 16 | 0 | true |
| wide-sparse | Sheet1!A1:Z40 | 384 | 866 | 851 | 0.983 | false | 3 | 4 | 0 | true |
| formula-heavy | Sheet1!A1:F60 | 512 | 1125 | 494 | 0.439 | true | 7 | 4 | 0 | true |
| metadata-heavy | Sheet1!A1:F24 | 448 | 1560 | 346 | 0.222 | true | 8 | 12 | 0 | true |
| public-formula-stress | Finance!A1:L62 | 640 | 1705 | 658 | 0.386 | false | 8 | 43 | 12 | true |

Recovery proof:

| Case | Same-range recovery | Locator metadata | Narrow sample recovery | Formula example recovery | Boundary |
| --- | --- | --- | --- | --- | --- |
| dense-table | true | true | A1:H8 | n/a | Use omitted locators or unbudgeted view |
| wide-sparse | true | true | A1:Z40 | n/a | Structural floor can exceed tiny requested budgets |
| formula-heavy | true | true | A2:F8 | n/a | Use omitted locators or unbudgeted view |
| metadata-heavy | true | true | A1:F8 | n/a | Metadata signals remain separate from raw evidence |
| public-formula-stress | true | true | A1:L9 | D48 | Formula-pattern examples recover representative evidence, not every omitted occurrence |

Decision: token-bounded agent view is now product-proof backed for deterministic summaries, omission metadata, locator recovery, and cross-surface budget metadata. Keep exact-token language out of release copy; wide sparse ranges can exceed very small requested budgets because column summaries are a structural floor.

## Product Example Gate

Token-bounded agent view is strong enough for a product example, but not for promotion into the top-two release proof index.

Allowed product example:

1. Request an agent view with a token budget.
2. Show `requestedApproxTokens`, `estimatedApproxTokens`, `unbudgetedApproxTokens`, `truncated`, and omission counters.
3. Use `budget.omittedEvidence.sampleRows` or `columnSamples` to choose a narrower follow-up range.
4. For formula omissions, use the formula-pattern example ref as representative recovery evidence.
5. State that same-range unbudgeted agent view is the exact fallback when locators are too broad.

Do not claim:

- exact token adherence;
- automatic recovery of every omitted formula occurrence;
- release-index parity with safe-open and package-action proof artifacts;
- that agent view replaces package inspection, raw reads, or mutation proof bundles.

## Next Product Prompt

```text
/goal Package Ascend's agent context contracts without adding new protocol surfaces. Document and test examples for token-bounded agent view and retained viewport patch history across existing SDK/CLI/API/MCP surfaces. Show budget omission counters, compact read changeToken polling, invalid/expired token recovery, and interactive viewport invalidation. Keep boundaries explicit: approximate token estimates, bounded retention, no collaborative CRDT claim, no unbounded history.
```

## Do Not Promote Yet

- CLI `read changedSince`: not exposed; keep the claim to SDK/API/MCP compact reads.
- Retention tuning knobs: wait for telemetry before expanding configuration.
- Collaborative editing language: this is patch history for one client workflow, not multi-user convergence.
- Exact token guarantees: the estimator is JSON bytes divided by four.
- Agent-view release proof index membership: keep the digest index limited to safe-open and package-action until product commits a persisted artifact policy for context-contract examples.
