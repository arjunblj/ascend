# Safe Open Batched Latency Probe

## Question

Can batching many `inspectWorkbookOpenPlan` calls per timing sample reduce safe-open open-plan CV enough to satisfy the owner-review profile?

## Hypothesis

If the current noisy latency evidence is dominated by timer overhead for sub-millisecond open-plan calls, then measuring each sample as a batch average should reduce CV across the required public cases. If batching does not reduce CV reliably, the release-latency blocker should stay blocked and the next performance loop should not spend implementation time on a simple batching flag.

## External sources checked

- Criterion.rs timing loops and batching guidance: https://bheisler.github.io/criterion.rs/book/user_guide/timing_loops.html
- Criterion.rs command output and sample/warmup model: https://bheisler.github.io/criterion.rs/book/user_guide/command_line_output.html
- Google Benchmark repeated statistics and coefficient of variation: https://github.com/google/benchmark/blob/main/docs/user_guide.md
- hyperfine warmup, exact runs, and JSON export options: https://man.archlinux.org/man/hyperfine.1.en

## Why this matters to Ascend

The QSS-leapfrog release matrix should not turn safe-open into a weak speed claim. The previous profile fold-in made noisy timing rejectable; this probe asks whether a tiny harness implementation would convert that blocker into acceptable evidence. If not, performance ownership should move to a better benchmark design rather than promoting local latency copy.

## Probe/implementation

Read `fixtures/benchmarks/safe-open-proof.ts` and confirmed the current harness measures each open-plan call independently with `performance.now()`, then computes per-case median, p95, and CV.

Ran two local throwaway probes over the six required public cases. Each probe:

- reused the same fixture bytes per case,
- performed 3 warmup samples and 10 measured samples,
- measured `batch` repeated `inspectWorkbookOpenPlan(bytes, { intent: 'edit-plan' })` calls inside one timer,
- reported per-call milliseconds by dividing elapsed time by `batch`.

No production code changed.

## Results

Batch size 50:

| Case | Median ms | CV |
| --- | ---: | ---: |
| clean | 0.605 | 0.37 |
| formula-heavy | 1.107 | 0.89 |
| macro | 0.331 | 0.32 |
| pivot | 1.071 | 0.43 |
| activex | 1.507 | 0.52 |
| chart | 1.035 | 0.24 |

Batch size 250:

| Case | Median ms | CV |
| --- | ---: | ---: |
| clean | 0.439 | 0.41 |
| formula-heavy | 1.171 | 0.22 |
| macro | 0.218 | 1.22 |
| pivot | 0.566 | 0.32 |
| activex | 1.158 | 0.73 |
| chart | 0.326 | 0.27 |

Batching alone does not satisfy the profile. It improved some cases but left multiple public cases above the CV <= 0.25 guard in both runs.

## Confidence

Medium. This was a local throwaway probe, not a committed harness. It is strong enough to reject a simple batching fold-in, but not enough to define the final safe-open performance benchmark.

## Fold-in decision

Kill the simple batching flag as the next fold-in. Keep `release-latency-run` blocked. The next performance owner should design a dedicated benchmark that controls process isolation, sample ordering, GC pressure, and artifact capture, or explicitly downgrade safe-open to non-latency release wording.

## Next question

Should safe-open release evidence use a separate process-level benchmark runner with randomized case order and artifacted raw samples, or should the release matrix permanently drop latency from the top safe-open claim?
