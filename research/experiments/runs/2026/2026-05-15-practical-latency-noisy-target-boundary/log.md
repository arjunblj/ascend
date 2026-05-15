# Practical Latency Noisy Target Boundary

## Question

When the practical-latency contract finds an above-floor but noisy envelope, should Ascend discard it as unusable, or should it route it as a guarded performance-owner target?

## Hypothesis

Noisy above-floor envelopes should stay visible as candidates, but the required next action must be remeasure-and-profile before production code changes. Below-floor phases should remain guardrails. None of this should become release claim evidence.

## External sources checked

- Google Benchmark user guide: repetitions/statistics guidance, including aggregate measurements for repeated runs. <https://github.com/google/benchmark/blob/main/docs/user_guide.md>
- hyperfine manual: warmup and repeated benchmark run controls. <https://man.archlinux.org/man/hyperfine.1.en>
- Bun benchmarking docs: local benchmark command conventions. <https://bun.com/docs/project/benchmarking>
- Criterion.rs timing loops: benchmark loop measurement discipline. <https://bheisler.github.io/criterion.rs/book/user_guide/timing_loops.html>

## Why this matters to Ascend

Ascend should not overclaim performance from local noisy timings, but performance loops still need a concrete next target. Dropping noisy above-floor timings can hide the only user-visible envelope worth investigating, while treating them as release evidence would overclaim.

## Probe/implementation

- Changed `fixtures/benchmarks/practical-latency-contracts.ts` so `productionTarget` selects above-floor envelopes even when their CV marks them noisy.
- Added the guard text `Remeasure the exact envelope and run this profile before code changes` for noisy targets.
- Kept below-floor envelopes as guardrails rather than production targets.
- Added regression coverage in `fixtures/benchmarks/practical-latency-contracts.test.ts` for a noisy 110 ms prepared-plan envelope.
- Confirmed `fixtures/benchmarks/release-proof-index.ts` still excludes `practical-latency-contracts` from release proof artifacts.

## Results

- Commit: `e363bfbb test(benchmarks): choose guarded latency contract target`.
- `bun test fixtures/benchmarks/practical-latency-contracts.test.ts --timeout 30000`: passed, 6 tests.
- `bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract edit-verify --repeat 1 --warmup 0 --dry-run --json`: passed and emitted public/generated input provenance plus skipped dry-run steps.
- `bunx biome check fixtures/benchmarks/practical-latency-contracts.ts fixtures/benchmarks/practical-latency-contracts.test.ts`: passed.
- `bunx tsc --build`: passed.
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`: passed and still reports `headlineClaimsAllowed: false`, `implementationSurfacePromotionAllowed: false`, and `missingRequirementCount: 9`.
- `bun test fixtures/benchmarks/release-proof-index.test.ts packages/sdk/src/release-trust-matrix.test.ts --timeout 30000`: passed, 6 tests.
- `bun run test:changed`: passed, 5217 pass, 1 skip, 0 fail.

## Confidence

High for the routing behavior and release boundary. Medium for the exact 5 ms production tuning floor because it remains a policy guard, not a release threshold.

## Fold-in decision

Promote to performance loop only as an owner action selector. Do not promote practical-latency output to release proof, performance thresholds, SLA wording, or headline claim evidence until a tracked-clean standardized public-input run has product-approved wording.

## Next question

No new research surface. The claim steward should keep the top-two release matrix collapsed and hand off only safe-open latency owner approval plus package-action boundary/fixture decisions.
