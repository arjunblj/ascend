# 2026-05 Ranked Research Portfolio

Date: 2026-05-15

## Portfolio Rule

Research is no longer a broad sweep. Each direction below must earn its place by proving a product-shaped claim, naming the evidence gap, and accepting a kill criterion. The top unknowns for this block are formula rejection proof and retained viewport patch examples because they close known proof gaps without adding new surfaces.

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
| 6 | Release proof index | Ascend can package top claim proof artifacts by digest without fake attestation claims. | Trustworthy releases and auditability. | Digest index, stable-shape hashes, privacy/artifact policy, no signed-provenance language. | Kill publication if storage semantics imply tamper evidence without signing. | Product/release |
| 7 | Formula conformance/oracle routing | Ascend can route formula mismatches by class across static goldens, HyperFormula, LibreOffice, and Excel. | Correctness credibility. | Runnable corpus artifacts, mismatch classes, oracle adapters, skip/divergence counters. | Kill public compatibility claims if private corpora or cached values are required. | Correctness |
| 8 | Property-based journal laws | Ascend can prove inverse journal laws across operation families better than hand fixtures alone. | Trustworthy mutation planning. | Fast-check generators, exact/lossy matrix, shrinkable counterexamples, changed-test integration. | Kill broad law claims if generators only cover toy sheets. | Correctness |
| 9 | Columnar scan sidecars | Ascend can accelerate repeated table/range scans with disposable generation-keyed sidecars. | Real-world performance without replacing workbook truth. | Real workbook benchmarks, build/invalidation/memory costs, checksum parity. | Kill product promotion if build+invalidation cost erases repeated-scan gains. | Performance |
| 10 | Agent workflow observability | Ascend can explain agent workflows as ordered, recoverable tool traces instead of opaque CLI/API calls. | World-class agent DX and observability. | Workflow traces for open-plan/read/agent-view/plan/commit, recovery prompts, failure taxonomy. | Kill if traces duplicate logs without improving repair or audit decisions. | Product/DX |

## Top Unknowns Proven This Block

### 1. Formula Rejection-First Cross-Surface Proof

Unknown: the board said formula language-service primitives still needed cross-surface `renameTarget` refusal snapshots.

Proof produced:

- Added CLI/API/MCP tests that call existing formula-assist surfaces with `=Budget+Sales[Amount]` at the table-name cursor.
- Each surface now asserts `renameTarget.ok === false`, `reason === "workbook-context-required"`, and role `table-name-use`.
- No edit-producing rename was implemented.

Validation:

```bash
bun test apps/cli/src/cli.test.ts -t "formula assist returns formula IDE help"
bun test apps/api/src/server.test.ts -t "formula-assist exposes diagnostics"
bun test apps/mcp/src/index.test.ts -t "ascend.formula_assist exposes formula IDE helpers"
bunx biome check apps/cli/src/cli.test.ts apps/api/src/server.test.ts apps/mcp/src/index.test.ts
```

Commit: `be666996 test(sdk): prove formula rename refusals across surfaces`.

Decision: Formula intelligence moved from "missing cross-surface refusal snapshot" to "needs latency/corpus proof before stronger release copy." Rename remains killed until workbook-context ownership exists.

### 1b. Formula Assist Corpus/Latency Proof

Unknown: after cross-surface refusal snapshots, the claim still needed a corpus-backed proof that formula assist is fast and refuses unsafe targets over realistic formulas.

Proof produced:

- Added `fixtures/benchmarks/formula-assist-proof.ts`.
- Added `fixtures/benchmarks/formula-assist-proof.test.ts`.
- The harness discovers formulas from public POI/ClosedXML formula fixtures, samples 250 formulas by default for proof runs, and combines them with explicit rejection-first cases for LET shadowing, defined names, table names, table columns, external references, 3D references, spill references, and function tokens.

Latest local proof:

| Metric | Value |
| --- | ---: |
| Public formulas discovered | 1685 |
| Sampled formulas | 250 |
| Static edge cases | 10 |
| Parse OK formulas | 260 |
| Reference spans | 506 |
| Binding roles | 19 |
| Prepare-rename OK targets | 3 |
| `no-symbol-at-cursor` refusals | 40 |
| `workbook-context-required` refusals | 3 |
| `reference-target-not-renameable` refusals | 214 |
| Median assist latency | 0.0252 ms |
| P95 assist latency | 0.0531 ms |
| Max assist latency | 2.4206 ms |

Validation:

```bash
bun test fixtures/benchmarks/formula-assist-proof.test.ts
bun run fixtures/benchmarks/formula-assist-proof.ts --public-formula-limit 250
bunx biome check fixtures/benchmarks/formula-assist-proof.ts fixtures/benchmarks/formula-assist-proof.test.ts
```

Decision: Formula language-service primitives are now claimable as corpus-backed, rejection-first primitives. They are still not a top release implementation handoff, and rename remains killed until a correctness-owned workbook-context symbol planner exists.

### 2. Retained Viewport Patch Product Proof

Unknown: whether retained viewport patch history is product-example ready without implying general sync or CRDT collaboration.

Proof produced:

- Reran the tracked viewport proof harness.
- Reran SDK interactive, API compact `changedSince`, and MCP compact `changedSince` validations.

Latest proof:

| Case | Observed | Passed | Patch bytes | Boundary |
| --- | --- | --- | ---: | --- |
| retained-patch | patch:A1 | true | 315 | SDK interactive retained token |
| skipped-token-retained | patch:A1 | true | 315 | bounded retained history |
| invalid-token | base-token-invalid | true | 0 | caller uses returned snapshot |
| cross-session-token | base-snapshot-missing | true | 0 | tokens are not shared history |
| expired-history | base-token-expired | true | 0 | retention is bounded |
| projection-change | base-snapshot-missing | true | 0 | no projection-specific code yet |
| metadata-invalidation | viewport-invalidated | true | 0 | metadata edits force refresh |

Validation:

```bash
bun run fixtures/benchmarks/viewport-patch-proof.ts
bun test fixtures/benchmarks/viewport-patch-proof.test.ts
bun test packages/sdk/src/interactive-contract.test.ts -t "retained|viewport patch results expose invalidation|tokens from other sessions"
bun test apps/api/src/server.test.ts -t "compact changedSince"
bun test apps/mcp/src/index.test.ts -t "compact changedSince"
```

Decision: Product example is allowed for SDK retained patches plus API/MCP compact recovery. CLI remains excluded. CRDT/collaboration language remains killed.

## Next Proof Moves

1. Safe-open release proof packaging: keep using existing surfaces and public fixtures; do not add new opener surfaces.
2. Auditable package-part mutation proof packaging: keep using existing proof/journal surfaces; do not add mutation surfaces.
3. Property-based journal laws: rank up only if generators cover real workbook features, not just scalar cells.
