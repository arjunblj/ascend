# 2026-05 Ranked Research Portfolio

Date: 2026-05-15

## Portfolio Rule

Research is no longer a broad sweep. Each direction below must earn its place by proving a product-shaped claim, naming the evidence gap, and accepting a kill criterion. The top unknowns for this block are the two release-claim candidates: safe unknown workbook opening and auditable package-part mutation. Formula and viewport work stays in proof-backed stewardship, but it is not the next implementation handoff.

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

## Top Unknowns Proven This Block

### 1. Safe Unknown Workbook Opening

Unknown: whether the top product/performance claim has current proof from existing SDK/CLI/API/MCP open-plan surfaces, without adding another opener.

Proof rerun:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1
bun test fixtures/benchmarks/safe-open-proof.test.ts packages/sdk/src/open-plan.test.ts
```

Latest proof:

| Metric | Value |
| --- | ---: |
| Proof cases | 9 |
| OK cases | 8 |
| Malformed rejected | 1 |
| Review before hydration | 4 |
| Public fixture open-plan speedup range | 14.09x to 31.97x |
| Synthetic signed mode | metadata-only, review |
| Synthetic unknown-part mode | metadata-only, review |

Decision: keep safe unknown workbook opening as the top handoff. The allowed claim is "package-feature routing and review branch before hydration." The honest boundary remains: not malware scanning, sandboxing, active-content safety, or malformed-package recovery.

### 2. Auditable Package-Part Mutation

Unknown: whether the second release claim still proves every package action kind and aligns package proof with rollback-journal evidence.

Proof rerun:

```bash
bun run fixtures/benchmarks/package-action-proof.ts
bun test fixtures/benchmarks/package-action-proof.test.ts
```

Latest proof:

| Metric | Value |
| --- | ---: |
| Proof cases | 8 |
| Passthrough actions | 27 |
| Regenerate actions | 38 |
| Add actions | 3 |
| Drop actions | 3 |
| Error actions | 1 |
| Cases with source graph evidence | 8 |
| Cases with package-preservation journal issue | 8 |
| Unknown-part proof issues | 1 |

Decision: keep auditable package-part mutation as the second handoff. The allowed claim is local per-part accounting with `passthrough`, `regenerate`, `add`, `drop`, and `error`. The honest boundary remains: not signed provenance, SLSA, in-toto attestation, Excel recalc equivalence, or semantic understanding of every unsupported feature.

### Formula Intelligence Guardrail

The latest rejection-first proof rerun is not a handoff to implementation. It exists to keep formula intelligence from overclaiming rename:

```bash
bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250
```

| Metric | Value |
| --- | ---: |
| Public formulas discovered | 1685 |
| Sampled formulas | 1685 |
| Reference spans | 2322 |
| Binding roles | 25 |
| Prepare-rename OK targets | 3 |
| Prepare-rename refusals | 1692 |
| P95 assist latency | 0.0368 ms |

Decision: formula intelligence remains a rejection-first primitives claim. Do not implement rename; do not hand it to an implementation loop until workbook-context symbol ownership and operation-owned edits exist.

### Release Proof Index

The current digest index covers only the top two artifacts:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings
```

| Artifact | Stable shape SHA-256 | Summary |
| --- | --- | --- |
| safe-open-proof | `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178` | cases=9, ok=8, rejected=1, reviewBeforeHydration=4, malformedRejected=true |
| package-action-proof | `9abebf576651551f58e00ccf8469d099b2c06dacd48391fe581a24e51a1e0afd` | cases=8, passthrough=27, regenerate=38, add=3, drop=3, error=1, streamingProofCases=1 |

Readiness gates now fail closed for the top two artifacts. The JSON index includes `readiness.releaseGate=blocked-by-publication-policy`, `readiness.headlineClaimsAllowed=false`, `totalRequirementCount=9`, `missingRequirementCount=9`, and `satisfiedRequirementCount=0`. It also points to compact report commands for each top artifact without indexing compact report digests.

| Artifact | Missing readyWhen gates |
| --- | --- |
| safe-open-proof | public-edge-fixtures; release-latency-run; publication-boundary; compact-report-publication-policy |
| package-action-proof | edge-fixture-policy; provenance-boundary; unsupported-feature-boundary; streaming-matrix-boundary; compact-report-publication-policy |

Decision: formula intelligence, retained viewport patches, token-bounded agent view, property-based journal laws, and columnar sidecars stay out of the top release proof index for now.

## Current Proof Refresh

Timestamp: 2026-05-15T05:52:35Z local proof run.

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
| Public fixture cases | 2 |
| Generated workbook cases | 2 |
| Generated edge-package cases | 4 |
| Passthrough actions | 27 |
| Regenerate actions | 38 |
| Add actions | 3 |
| Drop actions | 3 |
| Error actions | 1 |
| Cases with source graph evidence | 8 |
| Cases with package-preservation journal issue | 8 |
| Representative streaming proof cases | 1 |

Decision: hand off to correctness/product only. The proof supports local per-part accounting and journal-linked package evidence. It does not support signed provenance, SLSA, in-toto, Excel recalculation equivalence, chart byte-passthrough, or full streaming matrix parity.

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

## Next Proof Moves

1. Product/performance handoff: publish the safe unknown workbook opening proof bundle from existing surfaces and public fixtures.
2. Correctness/product handoff: publish the auditable package-part mutation proof bundle from existing proof/journal surfaces.
3. Correctness follow-up: property-style journal laws now have a deterministic tracked harness with pre-seeded metadata and package-state replacements, explicit style/table-style lossy boundaries, and a claim-safe report mode. They may rank up only after shrinkable generation or a correctness owner accepts deterministic proof as sufficient for release evidence.

## Stewardship Update: Prepared Agent Commits

Claim placement: supporting evidence under trustworthy mutation planning and agent workflow observability, not a new top release claim.

Proof added:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "prepared agent plans reuse full workflow state|prepared agent plans expose rollback journal safety facts|prepared agent commits"
bun test apps/mcp/src/index.test.ts -t "prepared plan|planHandle|writePolicyCheckMs"
bun test apps/api/src/server.test.ts -t "prepared|plan and commit"
```

Allowed wording: "prepared value-edit commits can reuse the planning structural check under a formula-free, one-shot, hash-guarded contract."

Boundary: this does not support a broad cache claim, formula-workbook reuse, structural-edit reuse, collaboration/snapshot reuse, or release-proof promotion without real-workbook latency evidence.
