# Safe Open Latency Validation Evidence

## Question

Can the safe-open release-latency blocker be made machine-readable without turning a local timing run into a release claim?

## Hypothesis

Yes. The release proof index can report whether the current handoff includes timed evidence, which command performance should run, and which policy requirements remain missing. Local timings can be logged as diagnostics while keeping release and threshold claims disabled.

## External sources checked

- Bun benchmarking docs: https://bun.sh/docs/project/benchmarking
- hyperfine repository and benchmark options: https://github.com/sharkdp/hyperfine
- hyperfine manual: https://man.archlinux.org/man/hyperfine.1.en

## Why this matters to Ascend

Safe unknown workbook opening is the rank-1 product claim, but speed claims are high-risk if they depend on one developer machine, dirty state, private corpora, or threshold wording that performance has not approved. The release artifact needs to guide the owner run, not smuggle in local measurements.

## Probe/implementation

- Inspected `safe-open-proof.ts` and `release-proof-index.ts`.
- Ran a local diagnostic:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json
```

- Added `safeOpenLatencyValidationEvidence` to the release proof index and owner-handoff JSON.
- The default `--no-timings` owner handoff reports `timed-evidence-absent-owner-run-required`, `releaseClaimAllowed=false`, and `thresholdClaimAllowed=false`.
- Added regression assertions in `fixtures/benchmarks/release-proof-index.test.ts`.

## Results

Local diagnostic summary:

| Case | Kind | Open-plan median ms | Full-open median ms | Full/open-plan ratio |
| --- | --- | ---: | ---: | ---: |
| `clean` | public fixture | 0.203 | 2.592 | 12.74 |
| `formula-heavy` | public fixture | 0.207 | 9.972 | 48.15 |
| `macro` | public fixture | 0.096 | 1.610 | 16.71 |
| `pivot` | public fixture | 0.160 | 2.446 | 15.32 |
| `activex` | public fixture | 0.132 | 2.079 | 15.74 |
| `chart` | public fixture | 0.084 | 1.579 | 18.89 |
| `signed` | generated edge package | 0.045 | 0.101 | 2.25 |
| `unknown-part` | generated edge package | 0.037 | 0.083 | 2.25 |

Default owner-handoff evidence:

| Field | Value |
| --- | --- |
| Status | `timed-evidence-absent-owner-run-required` |
| Timed case count | `0` |
| Release claim allowed | `false` |
| Threshold claim allowed | `false` |
| Missing policy requirements | tracked-clean release environment; standardized public input set; approved repeat and warmup policy; non-threshold release wording |

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
```

Result: 4 tests passed.

## Confidence

High that the handoff now prevents accidental latency promotion. Medium on the diagnostic numbers because they are local machine data and are intentionally not release evidence.

## Fold-in decision

Promote to performance and release loops as owner-routing evidence only. Keep `release-latency-run` missing and do not publish speed, SLA, threshold, or ratio claims from this diagnostic.

## Next question

Can the release packaging audit be moved into the research experiment structure and turned into owner prompts without changing package surfaces prematurely?
