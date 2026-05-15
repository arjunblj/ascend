# Safe Open Timing Environment Provenance

## Question

Can the safe-open latency proof capture enough runtime and machine context for performance-owner review without turning local timings into a release claim?

## Hypothesis

Distribution fields are necessary but not sufficient. A performance owner also needs the runtime, CPU, platform, and memory context to judge whether a timed proof can be compared or repeated. Capturing that metadata only when timings are enabled improves provenance while keeping no-timing release proof artifacts stable.

## External sources checked

- Bun benchmarking docs: https://bun.com/docs/project/benchmarking
- hyperfine docs and JSON/warmup options: https://github.com/sharkdp/hyperfine
- hyperfine manual: https://man.archlinux.org/man/hyperfine.1.en
- Google Benchmark repeated statistics: https://github.com/google/benchmark/blob/main/docs/user_guide.md

## Why this matters to Ascend

The top safe-open claim is strongest when it stays proof-shaped: pre-hydration routing with honest boundaries. Latency can support the claim later, but only if performance owners can see where the timings came from. Without environment provenance, a local median can be mistaken for a portable threshold.

## Probe/implementation

Ran:

- `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json`

Then folded timing environment metadata into `fixtures/benchmarks/safe-open-proof.ts`. The field is emitted only when `includeTimings` is true:

- runtime
- Bun version
- Node compatibility version
- platform
- architecture
- CPU model
- CPU count
- total memory bytes
- boundary language forbidding release attestation or hardware-normalized claims

`fixtures/benchmarks/release-proof-index.ts` now exposes `timingEnvironmentCaptured` and optional `timingEnvironment` on `safeOpenLatencyValidationEvidence`.

## Results

The local timed run produced environment metadata:

- runtime: Bun
- Bun: 1.3.13
- Node compatibility: v24.3.0
- platform: darwin
- arch: arm64
- CPU: Apple M4 Max
- CPU count: 14
- memory: 38,654,705,664 bytes

The no-timing release proof index still reports `timingEnvironmentCaptured: false`, which is the correct default for deterministic owner-handoff runs. Timed safe-open proof reports include environment metadata and the existing p95/CV/sample-count fields.

Validation:

- `bunx biome check fixtures/benchmarks/safe-open-proof.ts fixtures/benchmarks/safe-open-proof.test.ts fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`
- `bun test fixtures/benchmarks/safe-open-proof.test.ts --timeout 30000`
- `bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`
- `bunx tsc --build`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`
- `bun run test:changed` (5,212 pass, 1 skip, 0 fail)

## Confidence

High that environment metadata improves owner review and prevents accidental release-threshold wording. Medium that the captured fields are sufficient for a first release-latency profile. Low that they are sufficient for a public performance claim; the blocker still needs a tracked-clean, owner-approved run profile.

## Fold-in decision

Fold into the safe-open proof harness and release proof index as performance-owner evidence. Keep `release-latency-run` blocked. Do not promote QSS-beating latency, SLA, or threshold language.

## Next question

Should the performance loop define a release-latency profile object with minimum repeats, allowed public inputs, CV guardrails, and required tracked-clean status before the gate can become satisfiable?
