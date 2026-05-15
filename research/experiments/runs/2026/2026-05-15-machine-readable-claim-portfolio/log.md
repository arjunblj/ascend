# Machine-readable Claim Portfolio

## Question

Can Ascend stop treating research as a broad topic sweep by making the ranked 10-claim portfolio part of the release proof artifact itself?

## Hypothesis

If `release-proof-index` exports a ranked `claimPortfolio`, owner loops can consume the claim wording, North Star link, evidence needed, kill criterion, and handoff owner without reading a synthesis document or inferring priorities from deferred notes.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- LSP prepareRename in the 3.17 specification: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Model Context Protocol specification: https://modelcontextprotocol.io/specification/2024-11-05/index
- Apache Arrow columnar format: https://arrow.apache.org/docs/format/Columnar.html
- DuckDB Excel import: https://duckdb.org/docs/stable/guides/file_formats/excel_import
- HyperFormula named expressions: https://hyperformula.handsontable.com/guide/named-expressions.html
- SLSA source provenance: https://slsa.dev/spec/v1.2/source-requirements
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

The North Star needs research to convert into proof and owner routing. A portfolio only in Markdown can drift; a portfolio emitted by the proof index can be tested, diffed, and handed to correctness, performance, product, and release loops as concrete evidence.

## Probe/implementation

Folded a `claimPortfolio` array into `fixtures/benchmarks/release-proof-index.ts` and the owner-handoff JSON. The portfolio ranks 10 product-shaped directions:

1. safe unknown workbook opening
2. auditable package-part mutation
3. formula language-service primitives
4. token-bounded agent view
5. retained viewport patch history
6. release proof bundle
7. formula oracle routing
8. property-style journal laws
9. columnar scan sidecars
10. agent workflow observability

Each row carries status, evidence needed, kill criterion, likely handoff owner, handoff decision, optional proof command, and boundary. The top two rows are the only `top-implementation-handoff` rows; lower-ranked rows are proof-packaging-only or do-not-promote-yet.

## Results

Validation command:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json | jq '{portfolio: [.claimPortfolio[] | {rank, name, status, handoffDecision, owners: .likelyHandoffOwner}], handoffs: [.implementationHandoffs[] | {rank, artifact, claim, blockers: .blockingRequirementIds}], promotion: .implementationSurfacePromotionAllowed}'
```

Observed result:

- `claimPortfolio` has 10 ranked rows.
- Rows 1 and 2 are `top-implementation-handoff`.
- Rows 3 through 6 are `proof-packaging-only`.
- Rows 7 through 10 are `do-not-promote-yet`.
- `implementationSurfacePromotionAllowed=false`.
- Implementation handoffs remain only `safe-open-proof` and `package-action-proof`.

Targeted validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
git diff --check -- fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
```

All passed after formatting.

## Confidence

High that this improves portfolio stewardship without adding a product surface. Medium that the exact rank order will hold; the artifact makes rank changes explicit and reviewable.

## Fold-in decision

Promote to topic synthesis and release-owner proof routing. Do not promote any new SDK, CLI, API, or MCP surface from this cycle.

## Next question

Can the release packaging audit be moved from an untracked root note into the experiment structure and turned into owner prompts without changing package surfaces prematurely?
