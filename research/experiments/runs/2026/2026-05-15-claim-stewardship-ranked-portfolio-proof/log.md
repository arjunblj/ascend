# Claim Stewardship Ranked Portfolio Proof

## Question

Can Ascend stop broad research sweeping and hold a ranked portfolio of 8-12 product-shaped claims, while proving only the top one or two highest-leverage unknowns from existing artifacts and keeping formula rename frozen?

## Hypothesis

Yes. The current proof artifacts should keep safe unknown workbook opening and auditable package-part mutation as the only owner handoffs. Formula intelligence can remain proof-backed but rejection-first, with no edit-producing rename and no new implementation surface.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- LSP 3.17 `prepareRename`: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft LET function: https://support.microsoft.com/en-gb/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
- Microsoft names in formulas: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft structured references: https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- openpyxl tutorial preservation caveat: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS CE write options: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

Ascend's North Star needs claims that are credible to users and useful to owner loops. A ranked claim portfolio prevents research from turning every promising benchmark or primitive into a product surface. The right output is a proof ladder: what can be said today, what needs one owner decision or validation run, and what must stay speculative.

## Probe/implementation

No production code was changed. The probe reran the current top proof artifacts and refreshed the synthesis docs:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json
```

Updated:

- `research/experiments/syntheses/2026-05-ranked-research-portfolio.md`
- `research/experiments/syntheses/2026-05-release-claim-board.md`

## Results

Safe unknown workbook opening remains the top handoff:

- 9 proof cases.
- 6 public fixtures.
- 2 generated edge packages.
- 1 malformed package.
- 8 OK, 1 rejected.
- 4 review-before-hydration routes.
- Stable shape SHA-256: `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`.

Auditable package-part mutation remains the second handoff:

- 8 proof cases.
- 3 public fixtures.
- 2 generated workbooks.
- 3 generated edge packages.
- Action totals: `passthrough=32`, `regenerate=39`, `add=3`, `drop=3`, `error=1`.
- One representative streaming proof.
- Stable shape SHA-256: `0f9eb22498bc528a63adc40e59a6acbbe07022fde6b2414fcbee73b8b3a56e41`.

The release gate remains fail-closed:

- `headlineClaimsAllowed=false`.
- `implementationSurfacePromotionAllowed=false`.
- `releaseGate=blocked-by-publication-policy`.
- `missingRequirementCount=9`.
- Missing owners: product 2, correctness 1, performance 2, release 4.

Formula intelligence remains rejection-first:

- 1685 public formulas sampled.
- 2322 reference spans.
- 25 binding roles.
- 3 LET-local prepare-rename OK targets.
- 1692 prepare-rename refusals.
- `workbook-context-required=4`.
- `reference-target-not-renameable=1403`.
- No latency wording promoted because this proof used `--no-timings`.

## Confidence

High that the top two handoffs are correct for the next loop because both are backed by tracked proof harnesses and the release index machine-gates new surface promotion. Medium that release wording can be approved without more work because generated edge-package policy, timing policy, provenance boundaries, unsupported-feature wording, streaming boundary wording, and compact report publication policy all still need owner decisions.

## Fold-in decision

Promote to topic synthesis and owner handoff only. Do not add production surfaces. Do not implement formula rename. Hand off:

- Product/performance/release: safe unknown workbook opening proof packaging and owner approval.
- Correctness/product/performance/release: auditable package-part mutation proof packaging and owner approval.

Keep formula intelligence, token-bounded agent view, retained viewport patch history, columnar sidecars, formula oracle routing, and agent workflow observability out of implementation promotion for this block.

## Next question

Can the product and release owners explicitly accept disclosed generated structural fixtures for the top two claims, or must they acquire public binary replacements before any headline release wording is allowed?
