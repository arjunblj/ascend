# Claim portfolio proof refresh after fix

## Question

After finishing the in-flight verifier fix, do the ranked research portfolio and release-claim board still point to the same top one or two highest-leverage implementation handoffs?

## Hypothesis

Yes. The production fix improves correctness proof hygiene, but it does not change the portfolio ranking: safe unknown workbook opening and auditable package-part mutation remain the top handoffs, while formula intelligence remains rejection-first and non-promoted.

## External sources checked

- LSP 3.17 `prepareRename`, which supports refusing rename preparation rather than producing edits: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft LET function documentation, which frames LET names as formula-local variables: https://support.microsoft.com/en-au/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
- Open Packaging Conventions package/relationship/signature model: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft Protected View, used as the safe-open competitor boundary: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653

## Why this matters to Ascend

The research loop was producing useful surfaces faster than release claims could absorb them. The portfolio should act as a gate: prove the top product-shaped claims with current harnesses, hand off only the top owner loops, and keep speculative surfaces out of release wording.

## Probe/implementation

Reran current proof harnesses without adding production surfaces:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json
```

Updated the ranked portfolio synthesis with the fresh proof timestamp and current counts.

## Results

Safe unknown workbook opening:

- 9 cases.
- 8 OK cases.
- 1 malformed-package rejection.
- 6 public fixture cases.
- 2 generated edge-package cases.
- 4 review-before-hydration cases.

Auditable package-part mutation:

- 8 cases.
- Action counts: `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1`.
- Source graph evidence in all 8 cases.
- Package-preservation journal issues in all 8 cases.
- One representative streaming proof case.

Release proof index:

- 2 artifacts.
- `headlineClaimsAllowed=false`.
- `releaseGate=blocked-by-publication-policy`.
- 9 readyWhen requirements missing.
- Missing owner-loop counts: correctness 1, performance 2, product 2, release 4.

Formula assist refusal proof:

- 1685 public formulas discovered and sampled.
- 2322 reference spans.
- 25 binding roles.
- 3 formula-local LET prepare-rename OK targets.
- 1692 prepare-rename refusals.
- Refusal reasons: `no-symbol-at-cursor=285`, `workbook-context-required=4`, `reference-target-not-renameable=1403`.

## Confidence

High that the top-two ranking is still correct for this block. The release proof index fails closed and the proof harness counts stayed stable. High that formula intelligence should not be promoted to rename: the current proof is dominated by legitimate refusals and lacks workbook-context symbol ownership.

## Fold-in decision

Promote to topic synthesis only. Hand off:

1. safe unknown workbook opening to product/performance;
2. auditable package-part mutation to correctness/product.

Do not promote formula rename, columnar sidecars, retained patch history, or release provenance until their owner gates are closed.

## Next question

Should the next owner loop resolve the safe-open public-edge fixture gate by explicitly accepting disclosed generated package-topology fixtures, or by sourcing durable public signed/unknown-part binaries?
