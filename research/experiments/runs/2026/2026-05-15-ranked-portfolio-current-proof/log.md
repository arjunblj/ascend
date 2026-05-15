# Ranked Portfolio Current Proof

## Question

Can Ascend hold a ranked 8-12 direction research portfolio and prove only the top one or two highest-leverage unknowns, instead of broad-sweeping new research or adding another production surface?

## Hypothesis

Yes. The current proof artifacts should keep safe unknown workbook opening and auditable package-part mutation as the only top implementation handoffs. Formula intelligence should remain rejection-first and explicitly avoid edit-producing rename.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- LSP 3.17 `prepareRename`: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft LET function: https://support.microsoft.com/en-gb/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
- Microsoft structured references: https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- openpyxl tutorial preservation caveat: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS CE write options: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

Ascend's North Star needs release claims that can survive scrutiny: preservation-first XLSX opening, trustworthy mutation planning, real-world performance, formula intelligence, and world-class agent DX. A ranked portfolio should decide what to prove next and prevent research from turning every promising primitive into a product surface.

## Probe/implementation

No production code was changed. The probe reran the current proof artifacts:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json
```

Updated:

- `research/experiments/syntheses/2026-05-ranked-research-portfolio.md`
- `research/experiments/syntheses/2026-05-release-claim-board.md`
- `research/experiments/index.md`

## Results

Safe unknown workbook opening remains rank 1:

- 9 proof cases.
- 6 public fixtures.
- 2 generated edge packages.
- 1 malformed package.
- 8 OK, 1 rejected.
- 4 review-before-hydration routes.
- Stable shape SHA-256: `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`.

Auditable package-part mutation remains rank 2:

- 8 proof cases.
- 4 public fixtures.
- 2 generated workbooks.
- 2 generated edge packages.
- Action totals: `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1`.
- Stable shape SHA-256: `b4cf0755f2fd40ff577ab87b6600ee14a58bbd69608796c903759e7e8b4d29e8`.

The release gate remains fail-closed:

- `releaseGate=blocked-by-publication-policy`.
- `headlineClaimsAllowed=false`.
- `implementationSurfacePromotionAllowed=false`.
- `missingRequirementCount=9`.
- Implementation handoffs: `safe-open-proof`, `package-action-proof`.

Formula intelligence remains rejection-first:

- 1685 public formulas sampled.
- 2322 reference spans.
- 25 binding roles.
- 3 LET-local prepare-rename OK targets.
- 1692 prepare-rename refusals.
- No edit-producing rename surface.

Release packageability is stronger but still not a headline claim:

- SDK tarball smoke exists.
- CLI/API/MCP app tarball smoke exists.
- Installed CLI bin, installed API capabilities callback, and installed MCP capabilities callbacks are covered.
- Registry publication, artifact storage, provenance wording, API listener lifecycle, stdio MCP protocol session, and retention/privacy filtering remain release-owned blockers.

## Confidence

High that the ranked portfolio is current and that only the top two claims should be handed off. Medium that either claim is releasable without owner action, because generated edge-fixture policy, safe-open timing policy, unsupported-feature wording, streaming parity wording, provenance boundaries, and compact-report publication policy remain unresolved.

## Fold-in decision

Promote to topic synthesis and owner handoff only. Do not add production surfaces. Do not implement formula rename.

Top handoffs:

- Product/performance/release: safe unknown workbook opening proof packaging and owner approval.
- Correctness/product/performance/release: auditable package-part mutation proof packaging and owner approval.

Do not promote yet:

- Formula rename or workbook-context rename.
- Columnar scan sidecars.
- Formula oracle routing.
- Property-style journal-law product claims.
- Agent workflow observability claims.

## Next question

Can product/release owners approve disclosed generated structural fixtures for safe-open and package-action proof, or must the next proof loop acquire public binary replacements before any headline wording is allowed?
