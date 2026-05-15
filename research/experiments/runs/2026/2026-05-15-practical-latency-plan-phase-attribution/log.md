# Practical Latency Plan Phase Attribution

Date: 2026-05-15

## Question

When the edit/verify latency contract has both aggregate prepared-plan timing and shared plan phase timing, should the decision rank the aggregate prepared-plan span or the measured sub-phases?

## Hypothesis

If shared plan phase metrics are present, the practical latency decision should rank measured sub-phases such as `load-workbook` ahead of the aggregate prepared-plan timing and treat any leftover prepared-plan time as unassigned endpoint overhead. Otherwise, the report double-counts a large aggregate and hides the actionable phase.

## External sources checked

- Hyperfine benchmarking guidance, including warmups and statistical output: https://github.com/sharkdp/hyperfine
- Google Benchmark user guide, including benchmark result interpretation: https://google.github.io/benchmark/user_guide.html
- Bun profiling documentation: https://bun.sh/docs/runtime/debugger

## Why this matters to Ascend

Performance work should not chase an aggregate number that already contains measured sub-phases. Ascend's release claim board keeps practical latency diagnostic-only; this change improves owner routing by making the next profiled phase concrete without promoting latency thresholds.

## Probe/implementation

- `practical-latency-contracts.ts` now exports `practicalLatencyContractsTestHooks.envelopeDecisions` for focused tests without executing the benchmark on import.
- The runner now uses `if (import.meta.main) await run()`.
- When `sharedPlanPhaseMedianMs` exists, the edit/verify envelope ranks shared plan sub-phases and reports the aggregate prepared plan only as "Prepared plan unassigned endpoint overhead."
- Added focused tests for both paths: phase split available and phase split unavailable.

## Results

Proof commands run:

```bash
bun test fixtures/benchmarks/practical-latency-contracts.test.ts
bunx biome check fixtures/benchmarks/practical-latency-contracts.ts fixtures/benchmarks/practical-latency-contracts.test.ts
bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract edit-verify --repeat 1 --warmup 0 --timeout-ms 60000 --json
```

Observed evidence:

- Focused tests passed: 4 tests, 21 expectations.
- Public-tracked edit/verify dry run completed all three steps: workflow commit, post-write breakdown, and agent phase profile.
- The decision chose `Prepared reopen written output` at 72.247 ms as the largest phase, not aggregate prepared plan/open.
- The generated edit input and dirty worktree make the run diagnostic-only, not release-threshold evidence.

## Confidence

High for decision attribution logic under synthetic summaries. Medium for performance conclusions because the dry run used one sample and generated edit input.

## Fold-in decision

Promote to performance loop as a benchmark harness fix. Keep out of product/release claims: this improves phase routing, not user-facing latency wording.

## Next question

Can a performance owner run the same contract on a clean tree with approved repeat/warmup policy and decide whether prepared-output reopen or API first-view payload/open is the next real optimization target?
