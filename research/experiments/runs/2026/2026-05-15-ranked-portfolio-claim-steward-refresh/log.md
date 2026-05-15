# Ranked Portfolio Claim Steward Refresh

## Question

Can the ranked research portfolio be refreshed from current proof outputs so it remains a claim-selection artifact rather than another broad topic sweep?

## Hypothesis

Yes. The existing proof harnesses should still rank safe unknown workbook opening and auditable package-part mutation highest, while lower-ranked directions stay proof-backed hold or do-not-promote.

## External sources checked

- Microsoft Protected View documentation: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl formula parsing documentation: https://openpyxl.readthedocs.io/en/latest/formula.html
- SheetJS VBA blobs documentation: https://docs.sheetjs.com/docs/csf/features/vba/
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

The research portfolio should choose what Ascend proves next. If every promising idea becomes a surface, research competes with correctness, performance, and product loops. Claim stewardship keeps product-shaped proof ahead of implementation.

## Probe/implementation

Reran current proof commands:

- `bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json`
- `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json`

Updated the ranked portfolio and owner handoff synthesis with current package-action fixture provenance, action counts, stable shape digest, and the new correctness policy checklist note.

## Results

Safe-open remains rank 1:

- 9 cases
- 6 public fixtures
- 2 generated edge packages
- 1 malformed package
- 8 OK, 1 rejected
- 4 review-before-hydration routes
- stable shape `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`

Package-action remains rank 2:

- 8 cases
- 4 public fixtures
- 2 generated workbooks
- 2 generated edge packages
- action counts: `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1`
- one representative streaming proof
- stable shape `b4cf0755f2fd40ff577ab87b6600ee14a58bbd69608796c903759e7e8b4d29e8`

Release gate remains fail-closed:

- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- correctness policy has one pending `package-action-proof/unsupported-feature-boundary` checklist item

## Confidence

High for ranking the top two handoffs because the proof index now exposes the owner gates directly. Medium for the exact product wording because product, correctness, performance, and release owners still need to accept the unresolved gates.

## Fold-in decision

Research handoff only. No new SDK, CLI, API, MCP, formula rename, sidecar, or viewport surface should be promoted from this refresh. The top one or two implementation loops remain proof packaging and owner approval for safe-open and package-action claims.

## Next question

Should the next proof-producing loop resolve product acceptance for disclosed generated safe-open fixtures, or correctness acceptance for the package-action unsupported-feature boundary matrix?
