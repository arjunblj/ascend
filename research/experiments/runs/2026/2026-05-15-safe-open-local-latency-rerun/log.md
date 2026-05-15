# Safe Open Local Latency Rerun

## Question

Does the current safe-open proof harness produce useful latency evidence for the `release-latency-run` blocker without turning local measurements into a release performance claim?

## Hypothesis

Yes. A timed local run can show whether pre-hydration open-plan routing is materially cheaper than full workbook hydration across the current standardized proof cases, but it should not close the release blocker until a performance owner approves the environment, repeat count, inputs, and threshold wording.

## External sources checked

- Bun benchmarking docs describe using Bun timing APIs and treating benchmark methodology as part of the evidence: https://bun.com/docs/project/benchmarking
- hyperfine documents warmup runs and machine-readable benchmark export, reinforcing warmup/repeat discipline for command benchmarks: https://github.com/sharkdp/hyperfine
- Microsoft Protected View documents Office's trust-oriented review mode, which remains competitor contrast rather than Ascend's latency claim: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653

## Why this matters to Ascend

Safe unknown workbook opening is the top product-shaped claim. The product proof already shows routing and review decisions; the remaining performance question is whether package-feature inspection is plausibly cheap enough to use before full hydration on unknown files.

## Probe/implementation

Ran the existing proof harness with timings enabled:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 2 --json
```

The tracked worktree had no tracked edits before the run; longstanding untracked research and temporary folders remain outside the release proof.

## Results

Public fixture cases:

| Case | Bytes | Mode | Review | Open-plan median ms | Full-open median ms | Full/open-plan ratio |
| --- | ---: | --- | --- | ---: | ---: | ---: |
| clean | 9112 | formula | false | 0.403 | 2.474 | 6.13 |
| formula-heavy | 64769 | formula | false | 0.178 | 6.099 | 34.30 |
| macro | 12752 | metadata-only | true | 0.063 | 1.555 | 24.68 |
| pivot | 19460 | formula | false | 0.144 | 2.247 | 15.62 |
| activex | 12433 | metadata-only | true | 0.090 | 1.467 | 16.36 |
| chart | 10138 | formula | false | 0.072 | 1.097 | 15.31 |

Generated structural cases:

| Case | Bytes | Mode | Review | Open-plan median ms | Full-open median ms | Full/open-plan ratio |
| --- | ---: | --- | --- | ---: | ---: | ---: |
| signed | 2254 | metadata-only | true | 0.046 | 0.132 | 2.88 |
| unknown-part | 1697 | metadata-only | true | 0.041 | 0.073 | 1.78 |

Malformed input rejected before hydration with `Missing end of central directory record`.

## Confidence

Medium. The evidence is consistent with cheap pre-hydration routing on current proof fixtures, but it is local machine evidence over a small standardized set. It is not a release-environment benchmark and it does not establish public threshold wording.

## Fold-in decision

Promote to topic synthesis only. Keep `release-latency-run` missing in `release-proof-index` until the performance owner approves a release environment, public input set, repeat/warmup policy, and wording that avoids threshold overclaiming.

## Next question

Can the performance owner define the minimal release-environment safe-open latency acceptance policy, or should research keep the blocker as diagnostic-only and move back to package-action boundary evidence?
