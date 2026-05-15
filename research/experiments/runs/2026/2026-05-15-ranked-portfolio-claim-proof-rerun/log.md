# Ranked Portfolio Claim Proof Rerun

## Question

Can Ascend keep research as a ranked 8-12 direction portfolio, hand off only the top one or two product-shaped claims, and produce fresh proof for those claims without promoting another surface?

## Hypothesis

Yes. The top two highest-leverage unknowns should remain safe unknown workbook opening and auditable package-part mutation. A fresh proof rerun should either preserve those handoffs or expose a blocker that demotes them.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Language Server Protocol 3.17 rename and prepareRename: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- DuckDB Excel import documentation: https://duckdb.org/docs/stable/guides/file_formats/excel_import

## Why this matters to Ascend

Ascend's North Star needs proof-backed product claims, not a growing list of attractive surfaces. A ranked claim portfolio lets research say which claims are allowed today, which need owner proof packaging, and which should stay speculative until their evidence threshold is real.

## Probe/implementation

Reran the current top proof harnesses and owner handoff index without production changes:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
```

Updated the ranked portfolio and release claim board with the 2026-05-15T19:37:18Z proof refresh.

## Results

The portfolio remains exactly 10 directions:

1. safe unknown workbook opening: top implementation handoff
2. auditable package-part mutation: top implementation handoff
3. formula language-service primitives: proof-packaging only
4. token-bounded agent view: proof-packaging only
5. retained viewport patch history: proof-packaging only
6. release proof bundle: proof-packaging only
7. formula oracle routing: do not promote yet
8. property journal laws: do not promote yet
9. columnar scan sidecars: do not promote yet
10. agent workflow observability: do not promote yet

Safe-open proof: 9 cases, 8 OK, 1 malformed rejection, 6 public fixtures, 2 synthetic edge packages, 1 malformed package, and 4 review-before-hydration routes across macro, ActiveX, signature, and unknown-part risk families.

Package-action proof: 8 cases, 4 public fixtures, 2 generated workbooks, 2 generated edge packages, action totals `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1`, 8 cases with source graph evidence, 8 cases with package-preservation journal issues, and 1 representative streaming proof case.

Owner gate: `releaseGate=blocked-by-publication-policy`, `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, `missingRequirementCount=9`, and implementation handoffs remain exactly `safe-open-proof` and `package-action-proof`.

## Confidence

High that the ranked portfolio is current and that research should not promote formula rename, columnar sidecars, or new SDK/CLI/API/MCP surfaces during this block. Medium that ranks 7-10 are stable beyond the current release block because fresh correctness or performance evidence could reorder them later.

## Fold-in decision

Promote only to claim stewardship and owner-loop handoff. Do not fold in production code. The top two proof owners should package evidence and resolve owner gates; research should not add another product surface.

## Next question

Can the safe-open owner loop close `public-edge-fixtures` through explicit generated-fixture acceptance or approved public binary replacements without changing open-mode surfaces?
