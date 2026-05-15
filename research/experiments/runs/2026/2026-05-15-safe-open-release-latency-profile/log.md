# Safe Open Release Latency Profile

## Question

Can the `release-latency-run` blocker become a machine-checkable owner-review profile instead of an implicit request for "some timed run"?

## Hypothesis

If the release proof index declares the repeat count, warmup count, required public cases, required timing environment, required metrics, CV guard, and forbidden uses, then performance owners can close or reject the blocker from evidence rather than interpretation. A local timed run should remain diagnostic unless it satisfies the profile and owner wording is approved.

## External sources checked

- Google Benchmark repeated statistics and coefficient of variation: https://github.com/google/benchmark/blob/main/docs/user_guide.md
- hyperfine warmup, exact runs, min-runs, and JSON export options: https://man.archlinux.org/man/hyperfine.1.en
- hyperfine project docs: https://github.com/sharkdp/hyperfine
- Bun benchmarking docs: https://bun.com/docs/project/benchmarking

## Why this matters to Ascend

The top safe-open claim should not become a vague speed claim. Ascend can eventually use latency to support its trust/proof/runtime positioning, but only if the performance loop has an auditable profile for what counts. This keeps the QSS-leapfrog matrix from drifting into "faster than QSS" without a stable, public-input benchmark.

## Probe/implementation

Ran:

- `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json`

Folded `SAFE_OPEN_LATENCY_RUN_PROFILE` into `fixtures/benchmarks/release-proof-index.ts`:

- command: `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json`
- minimum repeat: 10
- minimum warmup: 3
- required public cases: clean, formula-heavy, macro, pivot, activex, chart
- required timing environment: true
- required metrics: sample counts, medians, p95, CV, and full/open ratio
- CV guard: public open-plan CV <= 0.25
- forbidden uses: release threshold, SLA, QSS performance comparison, hardware-normalized benchmark, private-corpus evidence, generated-only input evidence

The release proof index now reports `runProfile`, `runProfileSatisfied`, and `runProfileFailures` inside `safeOpenLatencyValidationEvidence`.

Validation:

- `bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`
- `bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`
- `bunx tsc --build`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`
- `bun run test:changed`

## Results

The local profile-shaped probe did not produce claimable release evidence. It satisfied repeat, warmup, and environment capture, but several public open-plan CVs exceeded the owner-review guard:

- `clean`: CV 0.39
- `formula-heavy`: CV 0.95
- `pivot`: CV 0.99

Other public cases were below the guard in this run:

- `macro`: CV 0.20
- `activex`: CV 0.13
- `chart`: CV 0.11

This is a useful guardrail result: the profile is specific enough to reject noisy local timing instead of encouraging weaker threshold wording. The no-timing owner handoff still reports `runProfileSatisfied: false` with concrete failures.

## Confidence

High that the profile makes the release-latency blocker more actionable. Medium that CV <= 0.25 is the right first guard; it is intentionally conservative and owner-review oriented. Low that any current local timing supports public latency wording.

## Fold-in decision

Fold into release proof owner handoff as a performance-loop blocker profile. Keep `release-latency-run` blocked. Do not promote QSS-beating latency, SLA, hardware-normalized, or threshold language.

## Next question

Should the performance loop improve the safe-open timing harness to reduce sub-millisecond measurement noise, for example by measuring batches of open-plan inspections per sample before recomputing CV?
