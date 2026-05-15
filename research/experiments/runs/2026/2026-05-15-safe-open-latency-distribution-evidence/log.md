# Safe Open Latency Distribution Evidence

## Question

Can the `release-latency-run` blocker for the safe unknown workbook opening claim be made more actionable without promoting a release speed claim?

## Hypothesis

The existing safe-open proof reports medians, but performance owners need distribution evidence to decide whether a release-environment run is stable enough for wording. Adding sample count, p95, and coefficient of variation will expose noise directly and prevent overclaiming from a single median.

## External sources checked

- hyperfine docs: https://docs.rs/hyperfine/1.9.0
- hyperfine project docs: https://github.com/sharkdp/hyperfine
- Google Benchmark user guide: https://github.com/google/benchmark/blob/main/docs/user_guide.md

## Why this matters to Ascend

The QSS-leapfrog matrix keeps safe-open as a top release claim only where Ascend can prove trust/proof/runtime behavior. A latency claim would be valuable, but only if the performance evidence is repeatable. Variance fields make the owner handoff more honest: low medians are visible, noisy runs stay blocked, and product copy cannot convert local observations into an SLA.

## Probe/implementation

Ran:

- `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 2 --json`

Then folded distribution fields into `fixtures/benchmarks/safe-open-proof.ts`:

- `openPlanSampleCount`
- `openPlanMedianMs`
- `openPlanP95Ms`
- `openPlanCv`
- `fullOpenSampleCount`
- `fullOpenMedianMs`
- `fullOpenP95Ms`
- `fullOpenCv`

Also surfaced public-file p95/CV maps in `fixtures/benchmarks/release-proof-index.ts` under `safeOpenLatencyValidationEvidence`. The release gate remains blocked and `releaseClaimAllowed` stays `false`.

## Results

The local probe showed public file open-plan medians below 1 ms, but several open-plan CVs were high enough to make threshold wording unsafe from this run alone:

- `clean`: open-plan median 0.326 ms, p95 0.400 ms, CV 0.24
- `formula-heavy`: open-plan median 0.339 ms, p95 2.670 ms, CV 1.19
- `macro`: open-plan median 0.117 ms, p95 0.147 ms, CV 0.10
- `pivot`: open-plan median 0.413 ms, p95 1.445 ms, CV 0.64
- `activex`: open-plan median 0.629 ms, p95 3.606 ms, CV 1.12
- `chart`: open-plan median 0.138 ms, p95 0.151 ms, CV 0.06

This is useful proof artifact data, not release performance proof. It strengthens the handoff by showing why the owner still needs a tracked-clean release-environment run over standardized public inputs.

Validation:

- `bunx biome check fixtures/benchmarks/safe-open-proof.ts fixtures/benchmarks/safe-open-proof.test.ts fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`
- `bun test fixtures/benchmarks/safe-open-proof.test.ts fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`
- `bunx tsc --build`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`
- `bun run test:changed` expanded to the full suite and failed under parallel 30s test limits: 5,198 pass, 1 skip, 14 fail, 4 errors.
- Isolated reruns of the full-suite failures passed with longer timeouts:
  - `bun test apps/cli/src/cli.test.ts --timeout 120000`
  - `bun test fixtures/benchmarks/competitive-io.test.ts fixtures/benchmarks/practical-latency-contracts.test.ts fixtures/benchmarks/prepared-plan-pressure.test.ts --timeout 120000`
  - `bun test fixtures/benchmarks/competitive-real-workbook.test.ts --timeout 120000`
  - `bun test packages/io-xlsx/src/writer/writer.test.ts --timeout 180000`
  - `bun test packages/sdk/src/agent-workflow.test.ts --timeout 120000`
  - `bun test fixtures/corpus/corpus.test.ts fixtures/xlsx/xlsx-fixtures.test.ts --timeout 120000`

## Confidence

High that the harness now exposes the evidence a performance owner needs to review variance. Medium that the local run is representative of this machine. Low for release wording because the run is local, unstored, and not owner-approved.

## Fold-in decision

Fold into the performance proof harness and release proof index. Keep `release-latency-run` blocked. Do not promote any QSS-beating latency, SLA, or threshold claim.

## Next question

Should the performance loop define an owner-approved latency run profile with minimum repeats, allowed public inputs, environment capture, and a CV guard before any release wording mentions speed?
