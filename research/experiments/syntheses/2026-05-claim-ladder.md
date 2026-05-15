# 2026-05 Product Claim Map

Date: 2026-05-15

## Purpose

Rank the product-shaped claims Ascend can prove, nearly prove, or should keep in research. This synthesis is deliberately promotion-limiting: it hands off only the top claims to implementation loops and treats the rest as proof backlog, not permission to add production surfaces.

## External Sources Checked

- [Microsoft Protected View](https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653) and [ProtectedViewWindow](https://learn.microsoft.com/cs-cz/office/vba/api/excel.protectedviewwindow): unsafe files are opened read-only and active content is restricted.
- [Microsoft workbook and macro signatures](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing): editing a signed workbook invalidates the signature.
- [openpyxl tutorial](https://openpyxl.readthedocs.io/en/stable/tutorial.html): `keep_vba` exists, but unsupported items such as shapes are not fully read and can be lost on save.
- [SheetJS write options](https://docs.sheetjs.com/docs/api/write-options/): writers focus on data preservation and features outside documented support may not serialize.
- [HyperFormula key concepts](https://hyperformula.handsontable.com/guide/key-concepts.html), [dependency graph](https://hyperformula.handsontable.com/guide/dependency-graph.html), and [function list](https://hyperformula.handsontable.com/docs/guide/built-in-functions.html): mature formula engines own AST reuse, dependency graphs, and broad function coverage.
- [Microsoft structured references](https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e) and [MS-OI29500 structured references](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/089fbdef-ed49-4a14-9509-794c95651b17): table names and column names are formula-language symbols, not plain text.
- [Univer MCP guide](https://docs.univer.ai/guides/sheets/getting-started/mcp) and [Univer MCP features](https://docs.univer.ai/guides/sheets/features/mcp): agent-native spreadsheet tools now expose MCP write/read/structure operations.
- [DuckDB Excel extension](https://duckdb.org/docs/lts/core_extensions/excel.html): `read_xlsx` turns sheets/ranges into typed SQL tables with inference boundaries.
- [Apache Arrow columnar format](https://arrow.apache.org/docs/format/Columnar.html): columnar memory optimizes scans and random access, with mutation coordination left to implementations.

## Claim Ladder

| Rank | Product claim | Status | What Ascend can honestly say now | What blocks stronger language |
| --- | --- | --- | --- | --- |
| 1 | Safe unknown workbook opening | Credible today, needs proof packaging | Ascend can inspect XLSX package features before full workbook hydration and recommend a load mode/review step across SDK, CLI, API, and MCP. | Needs a public proof bundle with fixture mix and latency numbers before it becomes a headline claim. |
| 2 | Auditable package-part mutation | Needs one more fold-in | Ascend already exposes package action proof and write-risk evidence for preserved/generated/skipped package content. | Needs stable product language for each package part: passthrough, regenerate, add, drop, or error, proven on real write workflows. |
| 3 | Token-bounded agent view | Credible today, needs product packaging | Ascend can emit deterministic workbook summaries with requested budgets, estimated budgets, and omission counters across agent-facing surfaces. | Needs a small product proof showing budget adherence and recovery from omitted evidence on diverse workbook shapes. |
| 4 | Retained viewport patch history | Credible in SDK, one product loop away | Ascend retains bounded change tokens for compact reads and interactive viewport patches with explicit invalidation reasons. | Needs product examples and telemetry/benchmark proof before implying general sync, history, or collaboration. |
| 5 | Formula language-service primitives | Credible primitives, not safe rename | Ascend exposes diagnostics, spans, hover, completions, reference edits, F4 cycling, and stateless binding roles. | Needs workbook-context binding resolution and `prepareRename`-style rejection before safe code-action claims. |
| 6 | Release proof bundle | Needs rank 1 and 2 | Ascend has local proof ingredients: inspect, plan, commit, reopen, diff, audit, digests, package action proof, and explicit boundaries. | Should wait until safe-open and package-part proof schemas stabilize; no signed provenance language yet. |
| 7 | Formula conformance and oracle routing | Correctness backlog | Ascend has mismatch-class and corpus routes in research and supporting tests. | Needs completed artifacts and runnable adapters for Excel, LibreOffice, HyperFormula, static goldens, and manual triage. |
| 8 | Columnar scan sidecars | Speculative product claim | Ascend has a performance research harness suggesting sidecars can accelerate repeated range scans. | Needs production invalidation semantics, real-workbook benchmarks, memory caps, and checksum parity before product promotion. |

## Proof Required By Claim

### 1. Safe unknown workbook opening

- Fixture: clean XLSX, formula-heavy XLSX, macro XLSM, signed workbook, pivot/slicer workbook, malformed package, and a workbook with unknown package parts. Public fixtures only.
- Benchmark: open-plan package fingerprint latency versus full hydration, with package size, part count, and chosen load mode recorded.
- API/CLI/MCP surface: existing `inspectWorkbookOpenPlan`, `ascend open-plan`, API open-plan endpoint, and MCP open-plan tool; do not add a new command unless the report cannot be generated from these.
- Validation gate: focused open-plan tests, malformed package tests, agent-doc workflow ordering, `bunx tsc --build`, `bunx biome check`, and `bun run test:changed` when implementation changes.
- Competitor contrast: Microsoft Protected View is trust/risk UX; Ascend's OSS angle is pre-hydration package-feature routing. openpyxl and SheetJS expose load/write options but do not make pre-hydration risk routing the first agent action.
- Honest boundary: not malware scanning, not sandboxing, not file trust, not proof that formula values or active content are safe.

### 2. Auditable package-part mutation

- Fixture: public XLSX/XLSM cases for document properties passthrough, worksheet XML regeneration, calc-chain drop, digital-signature invalidation, macro/ActiveX preservation, drawing/chart sidecar passthrough, and unknown part rejection.
- Benchmark: package proof overhead during plan and commit, reported as bytes and milliseconds against the same workflow without full proof expansion.
- API/CLI/MCP surface: full SDK `PackageActionProof`; compact CLI/API/MCP proof counts and optionally expanded part evidence with `--package-actions` or equivalent existing flags.
- Validation gate: plan/commit/reopen/diff/audit tests; package graph fidelity assertions; proof schema snapshot; targeted SDK/CLI/API/MCP tests; typecheck, Biome, changed tests for code changes.
- Competitor contrast: openpyxl warns unsupported shapes can be lost on save; SheetJS frames documented writer support around data preservation. Ascend's stronger claim is part-by-part write accounting, not just "we preserve data."
- Honest boundary: local evidence only, not cryptographic attestation, not SLSA, not assurance that Excel will semantically refresh every dependent feature.

### 3. Token-bounded agent view

- Fixture: wide sparse sheet, dense table sheet, formula sheet, metadata-heavy workbook, and mixed sheets with comments/styles/validations.
- Benchmark: emitted bytes and approximate token estimates against raw read, compact read, and full `agentView`; include omission counters.
- API/CLI/MCP surface: existing SDK `agentView`, CLI `agent-view`, API agent-view route, and MCP read/agent-view route with budget fields.
- Validation gate: deterministic truncation tests, cross-surface JSON shape tests, docs examples, and no hidden summarization without explicit omission metadata.
- Competitor contrast: Univer emphasizes agent operation routing through MCP; Ascend can differentiate on deterministic local workbook summaries under strict token budgets.
- Honest boundary: token counts are approximate; omitted evidence is intentionally not present; summaries cannot replace raw package inspection or proof artifacts.

### 4. Retained viewport patch history

- Fixture: unchanged viewport, changed cells, changed styles, changed metadata, stale token, invalid token, expired token, skipped intermediate token, changed projection/columns, and cross-session token.
- Benchmark: patch bytes and latency versus full viewport refresh under repeated small edits and periodic metadata invalidation.
- API/CLI/MCP surface: SDK compact reads and interactive session APIs today; product loop should show how API/MCP compact reads surface `changeToken` and invalidation metadata.
- Validation gate: compact changedSince tests, interactive viewport contract tests, invalidation reason assertions, and memory cap/retention tests.
- Competitor contrast: database MVCC and local-first systems retain older readable versions until compaction. Spreadsheet OSS usually exposes read/write snapshots rather than generation-aware patches.
- Honest boundary: bounded per-window retention, not unlimited history, not collaborative CRDT merge, not a transaction isolation guarantee across all workbook metadata.

### 5. Formula language-service primitives

- Fixture: LET names, defined names with workbook/sheet scope, table structured refs, sheet-qualified refs, external workbook refs, dynamic arrays, shared formulas, 3D refs, strings that look like refs, and parse failures.
- Benchmark: assist latency on large formula corpora and worst-case long formulas.
- API/CLI/MCP surface: existing SDK `formulaAssist` and `formulaBindingRoles`, CLI `formula assist`, API formula-assist endpoint, and MCP `ascend.formula_assist`.
- Validation gate: parser/token/span tests, binding-role tests, formula assist cross-surface tests, and oracle fixtures for ambiguous grammar.
- Competitor contrast: HyperFormula owns parser, AST reuse, dependency graph, and broad evaluation. Ascend should claim workbook-preserving edit intelligence, not formula-engine parity.
- Honest boundary: no safe rename yet; no guarantee that workbook-context names or external refs are resolved; code actions remain previews until `prepareRename` rejects ambiguous symbols.

### 6. Release proof bundle

- Fixture: one public workflow per claim: inspect/open-plan, plan, commit, reopen, diff, audit, and package proof over real XLSX/XLSM files.
- Benchmark: bundle generation overhead and artifact size compared with normal commit.
- API/CLI/MCP surface: stable SDK proof schema first; CLI/API report generation second; MCP can expose compact references after the schema is stable.
- Validation gate: golden proof bundle fixtures, digest checks, reopen/diff checks, package graph audit checks, and explicit failure cases.
- Competitor contrast: generic libraries can read/write files. Ascend's angle is an explainable decision trail from unknown input to verified output.
- Honest boundary: not signed provenance, not tamper-evident without external signing, not a public certification claim.

### 7. Formula conformance and oracle routing

- Fixture: static formula corpus plus routed mismatch classes: cached-value only, volatile, numeric drift, unsupported function, external refs, dynamic arrays, structured refs, date-system behavior, and semantic divergence.
- Benchmark: corpus completion time and per-oracle route overhead.
- API/CLI/MCP surface: completed JSON artifacts and a CLI report are enough; do not put this in MCP until the oracle outputs are stable.
- Validation gate: converter tests, completed artifact verifier, skipped/known-divergence counters, and no threshold changes without evidence.
- Competitor contrast: HyperFormula is the strongest OSS formula baseline; LibreOffice and Excel are external behavior oracles with their own automation boundaries.
- Honest boundary: cached workbook values are not truth; no blanket Excel-compatible formula claim.

### 8. Columnar scan sidecars

- Fixture: real workbook tables/ranges with numbers, dates, blanks, strings, formulas, filters, hidden rows, and style-heavy sheets.
- Benchmark: repeated scans, sidecar build cost, invalidation cost, memory overhead, and checksum parity against canonical workbook reads.
- API/CLI/MCP surface: performance benchmark harness only until proven; maybe private SDK cache later.
- Validation gate: generation-key invalidation tests, checksum parity, memory cap tests, and benchmark regression guard before production.
- Competitor contrast: DuckDB already reads XLSX ranges into typed SQL tables; Arrow defines the columnar scan substrate. Ascend's possible advantage is disposable sidecars over preservation-first workbook truth.
- Honest boundary: not a replacement storage engine, not a workbook rewrite, not guaranteed faster for sparse or single-pass reads.

## Status Buckets

### Credibly claim today

- Safe unknown workbook opening, phrased as "Ascend recommends a load mode and trust-review step from package features before full workbook hydration."
- Token-bounded agent view, phrased as "Ascend emits deterministic workbook summaries with requested budget metadata and omission counters."
- Retained viewport patch history, phrased as "Ascend can patch from bounded retained tokens and returns explicit invalidation reasons when it cannot."

### Needs one more fold-in

- Auditable package-part mutation: stabilize the part action proof vocabulary and fixture-backed workflow evidence.
- Formula language-service primitives: add workbook-context symbol resolution and `prepareRename` rejection before safe edits.
- Release proof bundle: package ranks 1 and 2 into a stable artifact after their schemas are settled.
- Formula conformance/oracle routing: make completed artifacts runnable and audited.

### Still speculative

- Columnar scan sidecars as a product claim.
- Collaborative or CRDT-like spreadsheet sync.
- Signed or third-party-attested release provenance.
- Universal Excel formula compatibility.
- Automatic trusted active-content execution or malware detection.

## Top Handoffs Only

### 1. Product/performance loop: safe unknown workbook opening proof bundle

```text
/goal Prove safe unknown workbook opening as a product claim. Use existing SDK/CLI/API/MCP open-plan surfaces; do not add a new user-facing surface. Build a fixture-backed report over public clean XLSX, formula-heavy XLSX, macro XLSM, signed workbook, pivot/slicer workbook, malformed package, and unknown-package-part cases. Include package fingerprint, recommended load mode, reviewBeforeHydration, risk feature families, package counts, and latency versus full hydration. Keep boundaries explicit: pre-hydration risk routing, not malware scanning or sandboxing. Validate focused tests, docs ordering, typecheck, Biome, and changed tests if code changes.
```

### 2. Correctness/product loop: auditable package-part mutation

```text
/goal Promote auditable package-part mutation into one stable proof schema. Start from existing package action proof surfaces and define per-part outcomes: passthrough, regenerate, add, drop, and error. Prove the schema on public workflows covering docProps passthrough, generated worksheet XML, calc-chain drop, signature invalidation, macro/ActiveX preservation, drawing/chart sidecars, and unknown part rejection. Expose full SDK JSON and compact CLI/API/MCP proof summaries through existing flags or response options. Validate with targeted tests, package graph fidelity checks, typecheck, Biome, and changed tests.
```

## Next-Loop Prompts

### Correctness

```text
/goal Build the auditable package-part mutation proof. Keep writer behavior unchanged unless a tiny evidence gap blocks the proof. Deliver a stable proof schema, fixture-backed tests, and compact cross-surface summaries for passthrough/regenerate/add/drop/error.
```

### Performance

```text
/goal Benchmark safe unknown workbook opening. Measure package-level open-plan latency versus full hydration across public clean, formula-heavy, macro, signed, pivot/slicer, malformed, and unknown-part workbooks. Add no production behavior unless instrumentation is missing and has a clear owner.
```

### Product

```text
/goal Package safe unknown workbook opening and token-bounded agent view as proof-first agent contracts. Use existing surfaces, write examples that show next safe actions and omission/review metadata, and avoid adding another narrow command unless the proof cannot be expressed otherwise.
```

## Do Not Promote Yet

- Columnar scan sidecars: keep in performance research until real-workbook benchmarks prove net benefit after build, invalidation, and memory costs.
- Safe formula rename: keep in research until workbook-context symbols and a rejection-first rename contract exist.
- Release proof bundle as provenance: do not imply signed, tamper-evident, SLSA, or third-party attestation.
- Universal formula compatibility: only claim routed oracle coverage by mismatch class.
- Collaborative spreadsheet sync: retained viewport history is bounded patch retention, not CRDT collaboration.
- Private workbook corpus results: never use private or large workbook data as public proof.
