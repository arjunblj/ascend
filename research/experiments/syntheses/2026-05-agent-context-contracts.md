# 2026-05 Agent Context Contracts

Date: 2026-05-14

## Product Claims

| Claim | Status | Proof now | Boundary |
| --- | --- | --- | --- |
| Token-bounded agent view | Product-ready after fold-in | SDK/CLI/API/MCP accept token budgets and return estimated tokens plus omission counters | Estimates are approximate and intentionally omit evidence-heavy samples |
| Retained viewport patch history | Credible, docs-ready | SDK/API/MCP compact reads and SDK interactive viewports return patches or explicit invalidation reasons | Bounded per-window/per-viewport history, not CRDT collaboration or unbounded event sourcing |

## Required Examples

| Example | Surface | Must show |
| --- | --- | --- |
| Budgeted sheet summary | CLI/API/MCP | requested tokens, estimated tokens, truncated flag, omitted sample rows, omitted column values |
| Stable compact poll | SDK/API/MCP read | first full compact window, unchanged poll, one-cell patch |
| Recovery from invalid token | SDK/API/MCP read | `base-token-invalid` and use returned full window |
| Recovery from expired token | SDK compact and SDK interactive | `base-token-expired` and use returned snapshot/window |
| Recovery from metadata/layout change | SDK interactive | `viewport-invalidated` and use returned snapshot |

## Next Product Prompt

```text
/goal Package Ascend's agent context contracts without adding new protocol surfaces. Document and test examples for token-bounded agent view and retained viewport patch history across existing SDK/CLI/API/MCP surfaces. Show budget omission counters, compact read changeToken polling, invalid/expired token recovery, and interactive viewport invalidation. Keep boundaries explicit: approximate token estimates, bounded retention, no collaborative CRDT claim, no unbounded history.
```

## Do Not Promote Yet

- CLI `read changedSince`: not exposed; keep the claim to SDK/API/MCP compact reads.
- Retention tuning knobs: wait for telemetry before expanding configuration.
- Collaborative editing language: this is patch history for one client workflow, not multi-user convergence.
- Exact token guarantees: the estimator is JSON bytes divided by four.
