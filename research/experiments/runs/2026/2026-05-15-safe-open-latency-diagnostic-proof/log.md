# Safe Open Latency Diagnostic Proof

## Question

Can the performance owner reduce the `safe-open-proof/release-latency-run` blocker with a timed public-input proof packet while keeping the result below release-threshold wording?

## Hypothesis

Yes for diagnostic evidence, no for release readiness. A local timed rerun should quantify open-plan versus full hydration on public fixtures, but the release gate should remain missing until performance approves environment, repeat/warmup policy, public input set, and non-threshold wording.

## External sources checked

- Bun benchmarking docs recommend choosing benchmark tools carefully and using appropriate timing/profiling methods: https://bun.com/docs/project/benchmarking
- hyperfine supports warmup runs and repeated command timing, reinforcing the need to name repeat/warmup policy: https://man.archlinux.org/man/hyperfine.1.en
- Google Benchmark user guide documents explicit benchmark output and repetition controls: https://google.github.io/benchmark/user_guide.html
- MDN performance budgets frame budgets as regression limits, not one-off local timing claims: https://developer.mozilla.org/en-US/docs/Web/Performance/Performance_budgets
- web.dev performance budgets recommend baselines and ongoing measurement rather than treating a single run as a product guarantee: https://web.dev/articles/performance-budgets-101

## Why this matters to Ascend

Safe unknown workbook opening is the rank-1 product/performance claim. The performance part of the claim needs real timing evidence, but local timing can easily become dishonest release copy if it is treated as an SLA or threshold. This probe should produce a proof packet and preserve the boundary.

## Probe/implementation

Ran:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json
```

No code changed.

## Results

Public fixture timing summary:

| Case | Open-plan median ms | Full-open median ms | Full/open ratio |
| --- | ---: | ---: | ---: |
| `clean` | 0.208 | 2.495 | 12.03x |
| `formula-heavy` | 0.220 | 7.216 | 32.75x |
| `macro` | 0.108 | 1.835 | 17.03x |
| `pivot` | 0.166 | 3.166 | 19.06x |
| `activex` | 0.104 | 1.519 | 14.57x |
| `chart` | 0.075 | 1.320 | 17.50x |

Generated structural cases are intentionally not release-public timing proof:

| Case | Open-plan median ms | Full-open median ms | Full/open ratio |
| --- | ---: | ---: | ---: |
| `signed` | 0.071 | 0.114 | 1.60x |
| `unknown-part` | 0.040 | 0.078 | 1.97x |

Malformed package handling stayed fail-closed: `open-plan rejected: Missing end of central directory record`.

Release claim status remains blocked:

- one local run is not a release environment;
- repeat 3/warmup 1 is diagnostic, not an approved benchmark policy;
- generated signed/unknown/malformed cases are not public real-workbook timing proof;
- no threshold, SLA, or headline speed wording should be promoted.

## Confidence

Medium-high for diagnostic routing: the command ran on tracked public fixture paths and gives concrete medians/ratios. Low for release wording: no owner-approved environment, repeat policy, threshold, or publication rule exists.

## Fold-in decision

Archive as diagnostic performance evidence and keep `release-latency-run` missing. Do not fold into production and do not update release thresholds.

## Next question

Can package-action performance reduce `streaming-matrix-boundary` with one additional generated add/drop/error streaming probe, or should the owner explicitly accept one representative dirty-sheet case as the narrow release boundary?
