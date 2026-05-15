# Safe open dirty timing rerun

## Question

Can the safe unknown workbook opening latency gate be advanced by rerunning the timed safe-open proof now?

## Hypothesis

No. A timed proof run is useful diagnostic evidence, but it must not advance the release latency gate if the tracked tree is dirty before the run.

## External sources checked

- Microsoft Protected View frames unsafe file opening as read-only review rather than proof of trust: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft OPC fundamentals describe package parts, relationships, and digital signatures, with signer/origin validation left to consumers: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl documents unsupported workbook object preservation boundaries: https://openpyxl.readthedocs.io/en/3.1.0/tutorial.html
- SheetJS write options document macro preservation options and writer scope: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

Safe unknown workbook opening is the top product/performance handoff. The performance gate should be strict: timing evidence may inform owners, but release wording must wait for a tracked-clean, approved environment and threshold policy.

## Probe/implementation

Checked tracked state and ran the existing proof harness:

```bash
git status --porcelain=v1 -uno
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

## Results

Tracked state before timing:

```text
 M packages/verify/src/checker.ts
 M packages/verify/src/verify.test.ts
```

Timed proof result:

| Case | Kind | Mode | Review | Open-plan median ms | Full-open median ms | Ratio |
| --- | --- | --- | --- | ---: | ---: | ---: |
| clean | public fixture | formula | false | 0.241 | 2.171 | 8.99 |
| formula-heavy | public fixture | formula | false | 0.202 | 7.902 | 39.05 |
| macro | public fixture | metadata-only | true | 0.087 | 1.763 | 20.33 |
| pivot | public fixture | formula | false | 0.156 | 3.039 | 19.47 |
| activex | public fixture | metadata-only | true | 0.111 | 2.109 | 18.97 |
| chart | public fixture | formula | false | 0.086 | 1.440 | 16.74 |
| signed | synthetic | metadata-only | true | 0.047 | 0.119 | 2.51 |
| unknown-part | synthetic | metadata-only | true | 0.042 | 0.082 | 1.95 |

Release proof index result:

- `headlineClaimsAllowed=false`.
- `releaseGate=blocked-by-publication-policy`.
- 9 readyWhen requirements still missing.
- `safe-open-proof/release-latency-run` remains missing.

## Confidence

High that the timing numbers are useful local diagnostics. High that they must not be treated as release proof because the tracked tree was dirty and the performance owner has not approved environment, repeat count, input set, or threshold wording.

## Fold-in decision

Archive as diagnostic-only evidence. Do not update release claim wording and do not mark any release proof index gate satisfied.

## Next question

After the verify fix is committed, rerun safe-open timing from `git status --porcelain=v1 -uno` returning empty, then record whether the result is still only diagnostic or ready for performance-owner review.
