# Formula Free Reopen Timing Probe

## Question

Does the formula-free no-value formula-mode reopen path produce enough local timing evidence to justify performance-loop follow-up?

## Hypothesis

Yes. A formula-mode reopen that proves formula absence without hydrating scalar cells should reduce post-write validation time on formula-free scalar workbooks, while preserving sheet dimension evidence.

## External sources checked

- Bun benchmarking and CPU profiling documentation: https://bun.sh/docs/project/benchmarking
- hyperfine benchmark tool: https://github.com/sharkdp/hyperfine
- mitata benchmark library: https://github.com/evanwashere/mitata

## Why this matters to Ascend

Prepared commits do a post-write reopen to prove the output workbook. If simple scalar edits can skip scalar value hydration for formula-free workbooks, Ascend can keep the trust loop cheaper without weakening correctness evidence.

## Probe/implementation

Ran a local generated-workbook timing probe in `/private/tmp/ascend-formula-free-reopen-probe`:

- created a 5000 x 20 formula-free scalar workbook through the SDK.
- saved it to XLSX.
- opened it 7 times with `mode: "formula", formulaModeHydrateValues: true`.
- opened it 7 times with `mode: "formula", formulaModeHydrateValues: false`.
- checked the scan-only path still preserved `A1:T5000` as the sheet dimension.

## Results

Local diagnostic timing:

| Mode | Median | Min | Max |
| --- | ---: | ---: | ---: |
| formula hydrate values | 9.143 ms | 8.586 ms | 20.131 ms |
| formula scan only | 4.899 ms | 4.291 ms | 7.764 ms |

Median speedup was 1.87x on 100,000 generated scalar cells.

## Confidence

Medium for direction, low for release wording. This is a local generated input with a small repeat count, not an approved benchmark environment or public-workbook latency claim.

## Fold-in decision

Promote to performance loop for a proper phase-profile benchmark. Keep out of release thresholds.

## Next question

Should the phase-profile benchmark expose formula-mode hydrate-values versus scan-only attribution for prepared post-write reopen?
