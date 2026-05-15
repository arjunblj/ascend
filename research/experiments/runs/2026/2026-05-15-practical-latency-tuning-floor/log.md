# Practical Latency Tuning Floor

## Question

Should practical latency contracts select tiny hot-cache phases as production optimization targets when their medians are below a few milliseconds or the measurement is noisy?

## Hypothesis

No. Sub-5 ms phases should remain guardrails unless they regress. The performance owner needs target selection to focus on user-visible envelopes where an implementation loop can plausibly move real latency, not micro-noise that creates churn.

## External sources checked

- Bun benchmark documentation: https://bun.sh/docs/project/benchmarking
- hyperfine README and warmup guidance: https://github.com/sharkdp/hyperfine
- Google Benchmark user guide: https://google.github.io/benchmark/user_guide.html
- web.dev performance budgets overview: https://web.dev/articles/performance-budgets-101

## Why this matters to Ascend

The North Star includes real-world performance, but production optimization loops should not chase noisy hot-cache fragments that are too small to matter in user-facing workflows. The practical-latency harness is owner-routing evidence; it should identify one meaningful target or explicitly decline to select one.

## Probe/implementation

Folded the in-flight benchmark change into `fixtures/benchmarks/practical-latency-contracts.ts`:

- added a 5 ms minimum production target phase floor;
- marked phases below the floor as guardrails via `nextAction`;
- excluded below-floor and noisy phases from `productionTarget`;
- exported `productionTarget` through test hooks;
- added a regression test proving a 2.2 ms noisy hot-cache first-paint phase does not become a production target.

## Results

Validation passed:

```bash
bun test fixtures/benchmarks/practical-latency-contracts.test.ts
bunx biome check fixtures/benchmarks/practical-latency-contracts.ts fixtures/benchmarks/practical-latency-contracts.test.ts
```

The focused benchmark test file passed 5 tests and 23 assertions.

Dry-run shape check:

```bash
bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract repeated-inspection --repeat 1 --warmup 0 --dry-run --json
```

The dry run remained non-measuring and produced no target decisions, as expected.

## Confidence

High that the decision logic now avoids below-floor/noisy production targets. Medium for the exact 5 ms threshold; it is a policy guardrail and should remain adjustable by the performance owner.

## Fold-in decision

Promote to performance-loop harness behavior. This does not promote any latency claim, threshold claim, or product copy.

## Next question

Can safe-open latency be rerun from a tracked-clean worktree and summarized as performance-owner evidence without satisfying the release-latency gate automatically?
