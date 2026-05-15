# Claim Stewardship Proof Refresh

## Question

Can Ascend stop broad research sweeping and maintain a ranked research portfolio plus release-claim board from current proof, while refusing to promote formula rename?

## Hypothesis

Yes. The highest-value work in this block is not another production surface; it is a proof refresh that ranks 8-12 research directions, preserves honest product wording, and hands off only the top two unknowns to implementation loops.

## External sources checked

- LSP 3.17 `prepareRename`: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft names in formulas: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft structured references: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

Ascend's North Star needs release claims that can survive scrutiny: preservation-first XLSX opening, auditable mutation planning, real-world performance, formula intelligence, and agent DX. A claim board prevents research from slipping into production just because a narrow surface is available.

## Probe/implementation

No production files were changed. The proof refresh ran existing harnesses:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1
bun run fixtures/benchmarks/package-action-proof.ts
bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250
```

Updated synthesis artifacts:

- `research/experiments/syntheses/2026-05-ranked-research-portfolio.md`
- `research/experiments/syntheses/2026-05-release-claim-board.md`

## Results

The ranked portfolio contains 10 directions with claim, North Star link, evidence needed, kill criterion, and likely handoff owner.

Top implementation handoffs remain:

1. Safe unknown workbook opening.
2. Auditable package-part mutation.

Fresh safe-open proof:

- 9 proof cases.
- 8 OK cases.
- 1 malformed package rejected.
- 4 review-before-hydration cases.
- Public fixture open-plan speedup range: 14.09x to 31.97x.

Fresh package-action proof:

- 8 proof cases.
- Commit actions: passthrough=27, regenerate=38, add=3, drop=3, error=1.
- 8 cases with source graph evidence.
- 8 cases with package-preservation journal issues.

Formula intelligence stays out of implementation handoff:

- 1685 public formulas discovered and sampled.
- 2322 reference spans.
- 25 binding roles.
- 3 LET-local prepare-rename OK targets.
- 1692 prepare-rename refusals.
- P95 assist latency: 0.0368 ms.

The release-claim board keeps formula intelligence rejection-first: no edit-producing rename, no defined-name/table rename, no external/sheet/range rename, and no workbook-context rename claim.

## Confidence

High for the top-two claim order and handoff boundaries because the proof harnesses passed locally and the wording is tied to existing surfaces. Medium for lower-ranked directions because their proof is intentionally held outside the release index.

## Fold-in decision

Promote to topic synthesis only. Hand off safe unknown workbook opening to product/performance and auditable package-part mutation to correctness/product. Do not promote formula rename or columnar sidecars.

## Next question

Can the product/performance loop turn the safe unknown workbook opening proof into a release artifact without adding a new open surface?
