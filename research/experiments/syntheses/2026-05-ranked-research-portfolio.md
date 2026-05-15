# 2026-05 Ranked Research Portfolio

Date: 2026-05-15

## Portfolio Rule

Research is no longer a broad sweep. Each direction below must earn its place by proving a product-shaped claim, naming the evidence gap, and accepting a kill criterion. The top unknowns for this block are the two release-claim candidates: safe unknown workbook opening and auditable package-part mutation. Formula and viewport work stays in proof-backed stewardship, but it is not the next implementation handoff.

## Latest Claim Steward Refresh

Proof timestamp: 2026-05-15T10:15:42Z.

The portfolio ranking still holds after the current proof run. The top two implementation handoffs remain safe unknown workbook opening and auditable package-part mutation. The release gate remains fail-closed, so the next step is owner approval or fixture replacement, not a new SDK/CLI/API/MCP surface.

| Artifact | Current proof | Blocking owner gates | Handoff decision |
| --- | --- | --- | --- |
| Safe unknown workbook opening | 9 cases: 6 public fixtures, 2 generated edge packages, 1 malformed package; 8 OK, 1 rejected, 4 review-before-hydration routes. Stable shape SHA-256: `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`. | `public-edge-fixtures`, `release-latency-run`, `publication-boundary`, `compact-report-publication-policy`. | Hand off to product/performance/release for proof packaging and boundary approval only. |
| Auditable package-part mutation | 8 cases: 3 public fixtures, 2 generated workbooks, 3 generated edge packages; action totals `passthrough=32`, `regenerate=39`, `add=3`, `drop=3`, `error=1`; source graph evidence everywhere; one representative streaming proof. Stable shape SHA-256: `0f9eb22498bc528a63adc40e59a6acbbe07022fde6b2414fcbee73b8b3a56e41`. | `edge-fixture-policy`, `provenance-boundary`, `unsupported-feature-boundary`, `streaming-matrix-boundary`, `compact-report-publication-policy`. | Hand off to correctness/product/performance/release for proof packaging and boundary approval only. |
| Release proof index | `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, `missingRequirementCount=9`, `signed=false`, `attestation=false`. | Product 2, correctness 1, performance 2, release 4. | Do not promote formula rename, agent view, viewport history, columnar sidecars, oracle routing, or agent traces into release scope this block. |

Fresh proof commands:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json
```

External fixture note: a constrained public-candidate probe found `node-projects/excelForge/src/test/Book 1.xlsx` as an unknown-part safe-open candidate (`preservedOther`, `metadata-only`, SHA-256 `9c5426fa71ff68cc7e40e19e02b5992daf91da5754ef643d2db2f89bd70bb122`). This does not close `public-edge-fixtures`: the candidate is not vendored, the source repository API reports no repository license, and the signed-workbook fixture gap remains.

## External Anchors

- [Microsoft Protected View](https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653) and [Open Packaging Conventions](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview) frame the safe-open/package proof boundary.
- [LSP 3.17](https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/) and [Excel structured references](https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e) frame formula intelligence as refusal-first until workbook context is proven.
- [Model Context Protocol](https://modelcontextprotocol.io/specification/2024-11-05/index), [Univer MCP](https://docs.univer.ai/guides/sheets/getting-started/mcp), and [Microsoft Graph Excel APIs](https://learn.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0) frame agent-context competition.
- [PostgreSQL MVCC](https://www.postgresql.org/docs/17/mvcc-intro.html), [SQLite isolation](https://www.sqlite.org/isolation.html), [RocksDB snapshots](https://github.com/facebook/rocksdb/wiki/Snapshot), and [Automerge concepts](https://automerge.org/docs/reference/concepts/) frame retained patch history and the CRDT boundary.
- [DuckDB Excel import](https://duckdb.org/docs/stable/guides/file_formats/excel_import) and [Apache Arrow columnar format](https://arrow.apache.org/docs/format/Columnar.html) frame columnar sidecars.
- [HyperFormula named expressions](https://hyperformula.handsontable.com/guide/named-expressions.html), [openpyxl formula parsing](https://openpyxl.readthedocs.io/en/latest/formula.html), and [Apache POI formula support](https://poi.apache.org/components/spreadsheet/formula.html) frame the formula-intelligence competitor boundary.
- [fast-check property-based testing](https://fast-check.dev/docs/introduction/what-is-property-based-testing/) frames journal-law evidence as generated-property proof with shrinking, not more hand-written fixture examples.
- [SLSA source provenance](https://slsa.dev/spec/v1.2/source-requirements) and [GitHub artifact attestations](https://docs.github.com/actions/concepts/security/artifact-attestations) frame why release proof bundles must not imply signed provenance or attestation.

## Ranked Portfolio

| Rank | Direction | Claim | North Star link | Evidence needed | Kill criterion | Handoff owner |
| ---: | --- | --- | --- | --- | --- | --- |
| 1 | Safe unknown workbook opening | Ascend can choose a cautious open mode before full hydration from package features. | Preservation-first XLSX and trustworthy agent workflows. | Public fixture matrix, latency versus full hydration, SDK/CLI/API/MCP proof, malformed rejection. | Kill release-headline copy if synthetic signed/unknown cases cannot be replaced or explicitly disclosed. | Product/performance |
| 2 | Auditable package-part mutation | Ascend can explain each write as passthrough/regenerate/add/drop/error with local proof. | Trustworthy mutation planning and preservation. | Per-part fixture coverage, journal compatibility, compact/full proof shape, unknown-part error evidence. | Kill stronger language if chart/signature/provenance boundaries are blurred. | Correctness/product |
| 3 | Formula rejection-first language service | Ascend can expose formula intelligence while refusing unsafe workbook-context rename targets. | Formula intelligence without unsafe mutation. | Cross-surface refusal snapshots, LET shadowing matrix, defined/table/external ref rejection, latency corpus. | Kill rename promotion until workbook-context symbol ownership and operation-owned edits exist. | Product/DX plus correctness |
| 4 | Retained viewport patch history | Ascend can patch from bounded retained tokens or return explicit refresh reasons. | Real-world performance and UI/agent efficiency. | SDK patch proof, API/MCP `changedSince` recovery, invalid/expired token examples, patch bytes. | Kill collaboration/sync wording unless multi-writer convergence is implemented and tested. | Product/performance |
| 5 | Token-bounded agent view | Ascend can summarize workbook intent under approximate token budgets with omitted-evidence recovery hints. | World-class agent DX. | Budget/recovery proofs, cross-surface metadata, product example using locators. | Kill exact-token wording if structural floors exceed tiny budgets. | Product/DX |
| 6 | Release proof index | Ascend can package top claim proof artifacts by digest and owner-loop readiness gates without fake attestation claims. | Trustworthy releases and auditability. | Digest index, stable-shape hashes, `readyWhen` gates, privacy/artifact policy, no signed-provenance language. | Kill publication if storage semantics imply tamper evidence without signing or if `readyWhen` requirements remain hidden from owner loops. | Product/release |
| 7 | Formula conformance/oracle routing | Ascend can route formula mismatches by class across static goldens, HyperFormula, LibreOffice, and Excel. | Correctness credibility. | Runnable corpus artifacts, mismatch classes, oracle adapters, skip/divergence counters. | Kill public compatibility claims if private corpora or cached values are required. | Correctness |
| 8 | Property-style journal laws | Ascend can prove selected inverse journal laws across generated operation sequences, package-state replacements, and explicit lossy metadata/style boundaries. | Trustworthy mutation planning. | fast-check shrinking, changed-test integration. | Kill broad law claims until generated coverage is shrinkable and style/table-style gaps stop being lossy or are excluded from exact wording. | Correctness |
| 9 | Columnar scan sidecars | Ascend can accelerate repeated table/range scans with disposable generation-keyed sidecars. | Real-world performance without replacing workbook truth. | Multi-source external public workbook benchmarks; first SEC external workbook parity now exists. | Kill product promotion if build+invalidation cost erases repeated-scan gains, checksum parity fails, or evidence stays limited to one external source plus generated/stress fixtures. | Performance |
| 10 | Agent workflow observability | Ascend can explain agent workflows as ordered, recoverable tool traces instead of opaque CLI/API calls. | World-class agent DX and observability. | Workflow traces for open-plan/read/agent-view/plan/commit, recovery prompts, failure taxonomy. | Kill if traces duplicate logs without improving repair or audit decisions. | Product/DX |

## Current Proof Refresh

Timestamp: 2026-05-15T10:15:42Z local proof run.

This refresh answers the latest portfolio question directly: the top one or two highest-leverage unknowns are still the two release-claim candidates, not another formula rename or sidecar surface.

### Top Unknown 1: Safe Unknown Workbook Opening

Proof command:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json
```

Current proof:

| Metric | Value |
| --- | ---: |
| Cases | 9 |
| OK cases | 8 |
| Rejected malformed cases | 1 |
| Public fixture cases | 6 |
| Generated edge cases | 2 |
| Review before hydration | 4 |
| Macro/ActiveX/signature/unknown routed to metadata-only review | 4 |

Decision: hand off to product/performance only. The proof supports cautious pre-hydration package-feature routing across existing surfaces. The proof does not support malware scanning, active-content safety, trust, sandboxing, signed provenance, or release latency wording.

### Top Unknown 2: Auditable Package-Part Mutation

Proof command:

```bash
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json
```

Current proof:

| Metric | Value |
| --- | ---: |
| Cases | 8 |
| Public fixture cases | 3 |
| Generated workbook cases | 2 |
| Generated edge-package cases | 3 |
| Passthrough actions | 32 |
| Regenerate actions | 39 |
| Add actions | 3 |
| Drop actions | 3 |
| Error actions | 1 |
| Cases with source graph evidence | 8 |
| Cases with package-preservation journal issue | 8 |
| Representative streaming proof cases | 1 |

Decision: hand off to correctness/product only. The proof supports local per-part accounting and journal-linked package evidence. It does not support signed provenance, SLSA, in-toto, Excel recalculation equivalence, chart byte-passthrough, or full streaming matrix parity.

### Top Unknowns Proven This Block

| Rank | Claim | Proof produced | Decision |
| ---: | --- | --- | --- |
| 1 | Safe unknown workbook opening | 9 proof cases; 6 public fixtures; 2 generated edge packages; 1 malformed package; 8 OK; 1 rejected; 4 review-before-hydration routes; stable shape `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`. | Hand off to product/performance/release for proof packaging and owner approval only. |
| 2 | Auditable package-part mutation | 8 proof cases; 3 public fixtures; 2 generated workbooks; 3 generated edge packages; action totals `passthrough=32`, `regenerate=39`, `add=3`, `drop=3`, `error=1`; stable shape `0f9eb22498bc528a63adc40e59a6acbbe07022fde6b2414fcbee73b8b3a56e41`. | Hand off to correctness/product/performance/release for proof packaging and owner approval only. |

Everything else is deliberately withheld from implementation handoff. Formula intelligence remains rank 3 as a rejection-first primitives claim, not a rename project.

### Release Gate Proof

Proof command:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

Current gate:

| Field | Value |
| --- | --- |
| Release gate | `blocked-by-publication-policy` |
| Headline claims allowed | `false` |
| Missing readyWhen requirements | 9 |
| Top owner action rank 10 | `package-action-proof/edge-fixture-policy` and `safe-open-proof/public-edge-fixtures` |
| Top owner action rank 20 | `package-action-proof/unsupported-feature-boundary` |
| Top owner action rank 30 | `safe-open-proof/release-latency-run` |

Decision: the portfolio should hand off the top two claims to owner loops for proof packaging and boundary approval. Research should not promote formula rename, columnar sidecars, or another release surface during this block.

### Formula Rejection Proof

Proof command:

```bash
bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json
```

Current proof:

| Metric | Value |
| --- | ---: |
| Public formulas discovered | 1685 |
| Sampled formulas | 1685 |
| Reference spans | 2322 |
| Binding roles | 25 |
| Prepare-rename OK targets | 3 |
| Prepare-rename refusals | 1692 |
| `workbook-context-required` refusals | 4 |
| `reference-target-not-renameable` refusals | 1403 |

Decision: keep formula intelligence at rank 3 as a rejection-first primitives claim only. The proof supports formula-local LET guard behavior and refusal classification. It does not support edit-producing rename, workbook-context defined-name rename, table-column rename, sheet/range rename, external-ref rename, or broader formula IDE claims.

### Provenance Boundary Audit

Proof commands:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
```

Current gate:

| Field | Value |
| --- | --- |
| Release index signed | `false` |
| Release index attestation | `false` |
| Package provenance gate | `provenance-boundary(missing,release)` |
| Compact package boundary | not signed provenance, SLSA, in-toto, or third-party attestation |
| Compact report embeds proof digests/artifact bytes | `false` |

Decision: keep release proof index at rank 6 and keep `provenance-boundary` missing. The top package-action claim can use local evidence wording only. Do not promote SLSA, in-toto, GitHub artifact attestation, Sigstore, signed provenance, or transparency-log language without a release/security owner implementing a real attestation pipeline.

### Streaming Matrix Boundary Audit

Proof command:

```bash
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json
```

Current gate:

| Field | Value |
| --- | --- |
| Total package-action cases | 8 |
| Cases with streaming proof | 1 |
| Streaming proof case | `docprops-passthrough` |
| Package-action classes covered overall | `passthrough`, `regenerate`, `add`, `drop`, `error` |
| Package-action classes covered by streaming proof | `passthrough`, `regenerate` |
| Non-streaming public fixtures | `macro-passthrough`, `chart-sidecar-accounting` |

Decision: keep `streaming-matrix-boundary` missing. The package-action claim can say one representative streaming dirty-sheet proof exists, but it cannot say streaming parity covers add/drop/error, public macro/chart fixtures, or every package-action scenario.

### Compact Report Publication Audit

Proof commands:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
```

Current gate:

| Field | Safe open | Package action |
| --- | --- | --- |
| Compact report command present | yes | yes |
| `compact-report-publication-policy` readyWhen present | yes | yes |
| Compact digest indexed | no | no |
| Compact report embeds workbook bytes | no | no |
| Compact JSON bytes | 3755 | 4258 |

Decision: keep compact report publication at rank 6 under release proof index, not as a new product claim. Compact report commands are useful local proof pointers; compact report digests remain do-not-promote until storage, privacy filtering, canonicalization, and verification expectations are owner-approved.

## Next Proof Moves

1. Product/performance handoff: publish the safe unknown workbook opening proof bundle from existing surfaces and public fixtures.
2. Correctness/product handoff: publish the auditable package-part mutation proof bundle from existing proof/journal surfaces.

Everything else is either proof-backed hold or do-not-promote-yet. The next loop should not add formula rename, agent-view, viewport, sidecar, oracle, or observability surfaces unless an owner explicitly changes the claim priority and proof gate.
