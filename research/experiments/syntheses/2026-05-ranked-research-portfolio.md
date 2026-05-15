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

Readiness gates now fail closed for the top two artifacts. The JSON index includes `readiness.releaseGate=blocked-by-publication-policy`, `readiness.headlineClaimsAllowed=false`, `totalRequirementCount=6`, `missingRequirementCount=6`, and `satisfiedRequirementCount=0`.

| Artifact | Missing readyWhen gates |
| --- | --- |
| safe-open-proof | public-edge-fixtures; release-latency-run; publication-boundary |
| package-action-proof | edge-fixture-policy; provenance-boundary; unsupported-feature-boundary |

Decision: formula intelligence, retained viewport patches, token-bounded agent view, property-based journal laws, and columnar sidecars stay out of the top release proof index for now.

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
