# Safe Open Latency Diagnostic Rerun

## Question

Does the current safe-open proof have enough timed evidence to close `safe-open-proof/release-latency-run`?

## Hypothesis

No. The existing harness can produce useful local timing diagnostics, but release latency wording still needs performance-owner approval for environment, repeat policy, public inputs, and threshold language.

## External sources checked

- Bun benchmarking documentation: https://bun.sh/docs/project/benchmarking
- hyperfine repository and usage docs: https://github.com/sharkdp/hyperfine
- hyperfine manual: https://man.archlinux.org/man/hyperfine.1.en

## Why this matters to Ascend

Safe unknown workbook opening is the top product-shaped claim. Latency evidence matters because the claim says Ascend can inspect package features before full hydration. The proof should not turn one local machine run into a release performance threshold.

## Probe/implementation

Ran the existing timed proof harness:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json
```

No production code changed. The probe records medians for open-plan package inspection and full workbook open on the current public and generated proof cases.

## Results

Public fixture cases:

| Case | Open-plan median ms | Full-open median ms | Full/open-plan ratio |
| --- | ---: | ---: | ---: |
| clean | 1.394 | 17.225 | 12.35 |
| formula-heavy | 2.514 | 55.850 | 22.21 |
| macro | 0.144 | 11.484 | 79.57 |
| pivot | 0.642 | 16.060 | 25.00 |
| activex | 1.042 | 24.328 | 23.34 |
| chart | 0.133 | 11.812 | 88.98 |

Generated structural cases:

| Case | Open-plan median ms | Full-open median ms | Full/open-plan ratio |
| --- | ---: | ---: | ---: |
| signed | 0.139 | 0.175 | 1.26 |
| unknown-part | 0.062 | 0.125 | 2.03 |

Malformed bytes were rejected during open-plan as expected.

## Confidence

Medium for diagnostic value: the run confirms pre-hydration package inspection is cheaper than full open for the public proof fixtures. Low for release performance wording: this is a local run, not an approved release environment or threshold policy.

## Fold-in decision

Keep as diagnostic research evidence and update the claim board. Do not mark `release-latency-run` satisfied. Do not publish latency thresholds or ratios as release claims until performance approves inputs, repeat/warmup policy, environment notes, and wording.

## Next question

Should performance define a standard release-latency environment and repeat policy, or should product first resolve generated fixture acceptance for the safe-open claim?
