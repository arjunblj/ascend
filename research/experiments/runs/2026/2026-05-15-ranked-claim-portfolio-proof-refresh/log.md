# Ranked Claim Portfolio Proof Refresh

## Question

Can Ascend's research portfolio be ranked into 8-12 product-shaped directions, then reduced to the top one or two proof handoffs without promoting another narrow production surface?

## Hypothesis

Yes. The highest-leverage unknowns are still safe unknown workbook opening and auditable package-part mutation because both already have existing surfaces, tracked proof harnesses, explicit owner gates, and release-shaped claim wording. Other directions should stay in proof-backed hold or do-not-promote status.

## External sources checked

- Microsoft Protected View frames unsafe Office files as a review/trust decision, not a package-feature proof claim: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Open Packaging Conventions frames workbook files as packages with parts and relationships, matching per-part mutation evidence: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- SheetJS CE write options document writer scope and data-preservation boundaries: https://docs.sheetjs.com/docs/api/write-options/
- openpyxl documents that unsupported Excel items such as shapes can be lost on open/save, supporting Ascend's narrower preservation-boundary contrast: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- LSP 3.17 keeps `prepareRename` separate from edit-producing rename, supporting formula refusal-first wording: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- MCP tools documentation supports structured tool results, but does not change the top claim ranking: https://modelcontextprotocol.io/specification/2025-06-18/server/tools
- DuckDB documents direct Excel `read_xlsx` ingestion, keeping columnar sidecars framed as analytics/performance research: https://duckdb.org/docs/current/guides/file_formats/excel_import.html
- Apache Arrow documents a language-independent columnar memory format, which is a sidecar substrate reference, not workbook truth: https://arrow.apache.org/docs/format/Columnar.html

## Why this matters to Ascend

The portfolio should prevent research from turning every interesting primitive into a release claim. Product-shaped claims need proof requirements, kill criteria, and owners. This keeps Ascend focused on preservation-first XLSX, trustworthy mutation planning, real-world performance, formula intelligence boundaries, and agent DX without overclaiming.

## Probe/implementation

Inspected the current ranked portfolio and release claim board, then reran the existing top proof artifacts without adding surfaces or touching production code:

```bash
git status --short --branch
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

Updated `research/experiments/syntheses/2026-05-ranked-research-portfolio.md` with a current claim-steward refresh.

## Results

Safe unknown workbook opening remains the top handoff:

- 9 proof cases.
- 6 public fixtures.
- 2 generated edge packages.
- 1 malformed package.
- 8 OK cases and 1 rejected malformed case.
- 4 review-before-hydration routes.
- Stable shape SHA-256: `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`.
- Remaining blockers: `public-edge-fixtures`, `release-latency-run`, `publication-boundary`, `compact-report-publication-policy`.

Auditable package-part mutation remains the second handoff:

- 8 proof cases.
- 2 public fixtures.
- 2 generated workbook cases.
- 4 generated edge-package cases.
- Action totals: `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1`.
- Source graph evidence appears in every case.
- One representative streaming proof exists.
- Stable shape SHA-256: `9abebf576651551f58e00ccf8469d099b2c06dacd48391fe581a24e51a1e0afd`.
- Remaining blockers: `edge-fixture-policy`, `provenance-boundary`, `unsupported-feature-boundary`, `streaming-matrix-boundary`, `compact-report-publication-policy`.

The release proof index keeps the gate closed:

- `headlineClaimsAllowed=false`.
- `implementationSurfacePromotionAllowed=false`.
- `missingRequirementCount=9`.
- `signed=false`.
- `attestation=false`.
- Missing requirements by owner: product 2, correctness 1, performance 2, release 4.

## Confidence

High that the top-two handoff ranking is correct for the current codebase. Medium for release readiness because every top claim still depends on owner acceptance or replacement of fixture/policy blockers.

## Fold-in decision

Promote to topic synthesis and owner handoff only. Do not add a new SDK, CLI, API, MCP, formula rename, agent-view, viewport, or sidecar surface from this block.

## Next question

Can product and release owners explicitly accept disclosed generated structural fixtures for safe-open and package-action proof, or should the next loop acquire durable public binary replacements before any headline wording changes?
