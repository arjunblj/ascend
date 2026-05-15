# Release Performance Policy Checklist

## Question

Can the performance-owned release blockers get the same approval-checklist treatment as fixture policy without promoting latency thresholds or streaming parity?

## Hypothesis

Yes. The release proof index can expose pending performance decisions with validation commands, acceptance evidence, and rejection conditions while keeping `release-latency-run` and `streaming-matrix-boundary` missing.

## External sources checked

- Bun benchmarking documentation: https://bun.sh/docs/project/benchmarking
- hyperfine repository documentation: https://github.com/sharkdp/hyperfine
- hyperfine manual page: https://man.archlinux.org/man/hyperfine.1.en

## Why this matters to Ascend

Performance evidence is high leverage for the safe-open and package-action claims, but careless wording can turn local observations into implied release thresholds or streaming parity. The owner handoff needs exact approval criteria before any performance language reaches release copy.

## Probe/implementation

Added `performancePolicy` to `fixtures/benchmarks/release-proof-index.ts` and the compact owner-handoff JSON. It contains two pending performance decisions:

1. `safe-open-proof/release-latency-run`;
2. `package-action-proof/streaming-matrix-boundary`.

Each checklist item records `status: pending-owner-decision`, validation command, acceptance evidence, and `rejectIf` text. The Markdown release proof index renders the checklist under `## Performance Policy`.

## Results

`bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json` now includes `performancePolicy.approvalChecklist`. The release gate remains blocked:

- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- both performance checklist items are `pending-owner-decision`

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check --write fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
```

## Confidence

High that the checklist improves owner routing without changing production code. Medium that the specific repeat/warmup policy is final; performance owners may choose a stronger release environment or repeat policy.

## Fold-in decision

Promote to performance/release proof packaging only. Do not mark `release-latency-run` or `streaming-matrix-boundary` satisfied.

## Next question

Can the correctness-owned `unsupported-feature-boundary` gate be turned into a machine-readable approval checklist with allowed and forbidden wording for signatures, calc chains, chart XML, macros/ActiveX, unknown parts, and streaming scope?
