# Ranked Portfolio Current Proof Refresh

## Question

Can Ascend's research portfolio be ranked as 8-12 product-shaped directions, then narrowed to the top one or two unknowns with fresh proof instead of more broad topic sweep?

## Hypothesis

Yes. The top unknowns should remain the release-claim candidates that already have surfaces and harnesses: safe unknown workbook opening and auditable package-part mutation. Formula intelligence, columnar sidecars, and viewport history should stay proof-bounded and out of implementation handoff for this block.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- LSP 3.17 prepare/rename separation: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- HyperFormula named expressions: https://hyperformula.handsontable.com/guide/named-expressions.html
- openpyxl formula parsing/tokenizer: https://openpyxl.readthedocs.io/en/latest/formula.html
- Apache POI formula support: https://poi.apache.org/components/spreadsheet/formula.html
- Univer MCP spreadsheet agent workflow: https://docs.univer.ai/guides/sheets/getting-started/mcp
- Notion MCP/developer platform: https://developers.notion.com/guides/mcp/overview
- DuckDB replacement scans and Arrow integration: https://duckdb.org/docs/stable/clients/c/replacement_scans.html and https://duckdb.org/2021/12/03/duck-arrow.html
- fast-check property-based testing: https://fast-check.dev/docs/introduction/what-is-property-based-testing/
- SLSA source provenance: https://slsa.dev/spec/v1.2/source-requirements
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

The North Star needs claim stewardship. A research portfolio should decide which product claims Ascend can credibly prove next, which claims need one more owner loop, and which attractive directions should be killed or held until their evidence threshold is met.

## Probe/implementation

- Inspected `research/experiments/syntheses/2026-05-ranked-research-portfolio.md`, the release claim board, the current experiment index, and existing proof harnesses.
- Finished the in-flight shared formula member range proof separately in production:
  - `9665f591 fix(io-xlsx): preserve shared member ranges`
- Reran current proof without adding new surfaces:
  - `bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json`
  - `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json`
  - `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json`
- Updated `research/experiments/syntheses/2026-05-ranked-research-portfolio.md` with current proof refresh and newer external anchors.

## Results

Ranked portfolio directions:

1. Safe unknown workbook opening.
2. Auditable package-part mutation.
3. Formula rejection-first language service.
4. Retained viewport patch history.
5. Token-bounded agent view.
6. Release proof index.
7. Formula conformance/oracle routing.
8. Property-style journal laws.
9. Columnar scan sidecars.
10. Agent workflow observability.

Safe-open proof: 9 cases, 8 OK, 1 malformed rejection, 6 public fixture cases, 2 generated edge cases, and 4 review-before-hydration cases. Macro, ActiveX, signature, and unknown-part inputs route to metadata-only review.

Package-action proof: 8 cases, 27 passthrough actions, 38 regenerate actions, 3 add actions, 3 drop actions, 1 error action, source graph evidence in every case, package-preservation journal issue in every case, and 1 representative streaming proof case.

Release proof index: still fail-closed with `releaseGate=blocked-by-publication-policy`, `headlineClaimsAllowed=false`, 9 missing `readyWhen` requirements, and top owner actions focused on fixture disclosure/replacement and claim-boundary approval.

## Confidence

High that the top two handoffs are still the right leverage points. Medium that rank 8 versus rank 9 will stay stable because journal laws and columnar sidecars can trade places after the next correctness/performance evidence refresh.

## Fold-in decision

Promote to topic synthesis and owner-loop handoff only. Do not promote new production surfaces. Hand off:

- Safe unknown workbook opening to product/performance for proof packaging and boundary approval.
- Auditable package-part mutation to correctness/product for fixture policy and unsupported-feature boundary approval.

Do not promote formula rename, columnar sidecars, or release proof attestation claims.

## Next question

Can the product/performance owner resolve the safe-open public-edge-fixture and release-latency gates without changing any open surfaces?
