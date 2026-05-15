# Latency Input Provenance Validation

## Question

Which diagnostic probe is most useful while promotion is throttled: latency input provenance validation, public fixture search, or compact report privacy review?

## Hypothesis

Latency input provenance is the most useful immediate diagnostic because it directly protects the safe-open performance gate from private-corpus overclaiming.

## External sources checked

- SPEC CPU 2017 run and reporting rules: https://www.spec.org/cpu2017/Docs/runrules.html
- Google Benchmark user guide: https://google.github.io/benchmark/user_guide.html
- hyperfine warmup/run options: https://man.archlinux.org/man/hyperfine.1.en
- Google SRE service level objectives and latency discussion: https://sre.google/sre-book/service-level-objectives/

## Why this matters to Ascend

Safe unknown workbook opening still needs performance-owner latency evidence. If latency reports use local/private workbooks, the report must say diagnostic-only even when tracked code is clean.

## Probe/implementation

Validated the committed `practical-latency-contracts` input provenance behavior with:

```bash
bun run fixtures/benchmarks/practical-latency-contracts.ts --dry-run --contract first-view --repeat 1 --warmup 0 --json --out-dir /tmp/ascend-practical-latency-provenance-clean
```

Inspected `/tmp/ascend-practical-latency-provenance-clean/summary.md`.

## Results

JSON result:

| Field | Value |
| --- | --- |
| Tracked code dirty | `false` |
| Untracked entries | 20 |
| Input workbook | `research/excel-corpus/NYC_311_SR_2010-2020-sample-1M.xlsx` |
| Input workbook tracked | `false` |
| Input workbook exists | `true` |
| Input workbook release-claimable | `false` |
| Edit workbook | `fixtures/xlsx/stress/dense-100k.xlsx` |
| Edit workbook tracked | `false` |
| Edit workbook exists | `true` |
| Edit workbook release-claimable | `false` |

Markdown result:

- Input provenance labels both workbooks as `local/private`.
- Worktree label says `tracked clean with local/private inputs; diagnostic only`.
- Guardrail says numbers remain diagnostic until rerun from tracked-clean code with tracked benchmark inputs, or until local inputs are explicitly documented as private diagnostics.

## Confidence

High. This validates the reporting boundary but does not produce release latency numbers.

## Fold-in decision

Promote to topic synthesis only. The benchmark harness fix is already committed; this log records why it matters. Keep `safe-open-proof/release-latency-run` missing until a performance owner runs standardized public inputs and approves threshold wording.

## Next question

Can public fixture search be constrained to only the missing signed/unknown structural cases, avoiding another broad corpus sweep?
