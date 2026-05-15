# Latency Target Variance Routing

## Question

Can Ascend's practical latency contract choose implementation targets with enough tail and variance evidence to avoid optimizing noisy medians?

## Hypothesis

If `agent-first-window` emits per-phase p95 and coefficient-of-variation stats, `practical-latency-contracts` can label target stability and route noisy targets back to measurement/profile work before production changes.

## External sources checked

- Bun benchmarking documentation: https://bun.sh/docs/project/benchmarking
- hyperfine README and CLI options for warmups/runs/exported measurements: https://github.com/sharkdp/hyperfine
- hyperfine manual page documenting warmup/min-runs/max-runs behavior: https://man.archlinux.org/man/hyperfine.1.en

## Why this matters to Ascend

Performance work should improve real user-visible envelopes, not chase a single median from a noisy local run. The performance owner needs enough evidence to decide whether to profile, remeasure, or implement.

## Probe/implementation

Folded a small stats summary into `fixtures/benchmarks/agent-first-window.ts`:

- sample count
- min
- median
- mean
- p95
- max
- standard deviation
- coefficient of variation

The summary is emitted for API first-window, API warm first-window, capped warm open window, TUI first paint, and warm TUI first paint. `fixtures/benchmarks/agent-phase-profile.ts` also emits per-phase stats for shared plan/commit phases used by edit-verify routing. `fixtures/benchmarks/practical-latency-contracts.ts` now propagates p95/CV into first-view, edit-verify, and repeated-inspection target decisions, so existing stability logic can mark candidates `stable`, `noisy`, or `unknown`.

## Results

Focused tests:

```bash
bun test fixtures/benchmarks/agent-first-window.test.ts fixtures/benchmarks/practical-latency-contracts.test.ts
bun test fixtures/benchmarks/agent-phase-profile.test.ts
bunx biome check fixtures/benchmarks/agent-first-window.ts fixtures/benchmarks/agent-first-window.test.ts fixtures/benchmarks/practical-latency-contracts.ts fixtures/benchmarks/agent-phase-profile.ts fixtures/benchmarks/agent-phase-profile.test.ts
git diff --check -- fixtures/benchmarks/agent-first-window.ts fixtures/benchmarks/agent-first-window.test.ts fixtures/benchmarks/practical-latency-contracts.ts fixtures/benchmarks/agent-phase-profile.ts fixtures/benchmarks/agent-phase-profile.test.ts
```

Probe command:

```bash
bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract first-view --repeat 1 --warmup 0 --timeout-ms 60000 --json
```

Observed decision excerpt:

- largest phase: `API first-view payload/open`
- phase median: `17.596875 ms`
- p95: `17.596875 ms`
- CV: `0`
- stability: `stable`
- next action: `profile required before production changes`

The one-sample probe is not release evidence; it proves the artifact shape and routing fields only.

## Confidence

High that the JSON/Markdown decision path now carries p95/CV when available. Medium on the exact `CV > 0.1` stability threshold; it is a conservative routing guard and should be revisited with multi-run public-tracked data.

## Fold-in decision

Promote to performance loop. This is a measurement-proof refinement only. It does not authorize latency claims, benchmark thresholds, or production optimization work without an owner-approved profile.

## Next question

Should the performance owner add a checked-in public-tracked multi-run profile that exercises the same p95/CV path with enough samples to make the stability label meaningful?
