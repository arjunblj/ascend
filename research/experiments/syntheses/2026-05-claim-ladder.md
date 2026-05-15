# 2026-05 Claim Ladder

Date: 2026-05-14

## Purpose

Rank product-shaped claims Ascend can make or should prove next. This synthesis intentionally stops the research loop from promoting every nearby surface. A claim graduates only when the proof is fixture-backed, surfaced through the right product channel, validated by gates, contrasted with credible alternatives, and bounded honestly.

## External Contrast Checked

- HyperFormula dependency graph and AST pipeline: https://hyperformula.handsontable.com/guide/dependency-graph.html and https://hyperformula.handsontable.com/guide/key-concepts.html
- HyperFormula built-in function breadth: https://hyperformula.handsontable.com/docs/guide/built-in-functions.html
- Univer MCP spreadsheet agent workflow: https://docs.univer.ai/guides/sheets/getting-started/mcp
- DuckDB Excel extension and `read_xlsx` semantics: https://duckdb.org/docs/lts/core_extensions/excel.html
- openpyxl `keep_vba` and unsupported item loss warning: https://openpyxl.readthedocs.io/en/3.0/usage.html
- SheetJS CE write options and data-preservation framing: https://docs.sheetjs.com/docs/api/write-options/
- ExcelJS read/write/streaming positioning: https://exceljs.org/

## Ranked Claim Ladder

| Rank | Claim | Status | Why this rank | Proof required |
| --- | --- | --- | --- | --- |
| 1 | Safe unknown workbook opening | Credible today, needs a proof bundle before marketing | Open-plan now exists across SDK, CLI, API, MCP, and agent docs. It is product-shaped because it changes the first action on unknown XLSX/XLSM files. | Fixture: clean workbook, macro workbook, signature/active-content workbook, pivot/slicer workbook, malformed package. Benchmark: package fingerprint latency versus full hydration. Surface: `open-plan` in CLI/API/MCP plus SDK. Validation gate: focused open-plan tests, bundled-doc ordering harness, `bun run test:changed`. Competitor contrast: openpyxl and generic read/write libraries expose load options but do not frame pre-hydration risk as the first product step. Honest boundary: not malware scanning, not a sandbox, and not permission to read workbook text before trust review. |
| 2 | Auditable package-part mutation | Needs one more fold-in | Package action proof is exposed, but the product claim needs stable proof language over real edit workflows. This is the strongest preservation-first differentiator. | Fixture: real XLSX/XLSM fixtures with docProps passthrough, calc-chain drop, generated sheet XML, signature invalidation, macro/ActiveX preservation blocks. Benchmark: proof overhead on plan/commit. Surface: compact proof in CLI/API/MCP, full SDK result. Validation gate: plan/commit/reopen/diff/audit tests and package graph assertions. Competitor contrast: openpyxl documents that unsupported items such as images/charts can be lost on save; Ascend should prove each part as passthrough/regenerate/add/drop/error. Honest boundary: local evidence, not cryptographic attestation or SLSA provenance. |
| 3 | Token-bounded agent view | Credible today, but should be packaged as a claim | Agent-view budget metadata exists and directly serves agent DX. It is less risky than mutation claims and should be easy to prove. | Fixture: wide sheet, formula sheet, table sheet, sparse sheet, metadata-heavy workbook. Benchmark: emitted bytes and approximate tokens versus raw read/agent-view. Surface: SDK/CLI/API/MCP examples with requested and estimated budgets. Validation gate: deterministic truncation tests and docs snippets. Competitor contrast: Univer is agent-native through MCP, while Ascend can emphasize deterministic file-summary budgets for local workbook agents. Honest boundary: token estimates are approximate and summaries omit evidence by design with explicit counters. |
| 4 | Retained viewport patch history | Credible in SDK, one product loop away | Compact reads and interactive SDK sessions now retain bounded token history. The product claim needs a client-facing budget/telemetry story before broad promotion. | Fixture: stable viewport, skipped token, expired token, metadata invalidation, cross-session token. Benchmark: patch bytes and latency versus full viewport refresh. Surface: SDK interactive session today; compact read through SDK/API/MCP; future docs should separate these. Validation gate: compact changedSince tests, interactive contract tests, `bunx tsc --build`, `bun run test:changed`. Competitor contrast: database snapshots and local-first sync systems retain history with explicit expiry; spreadsheet OSS typically exposes reads/writes rather than generation-aware patch streams. Honest boundary: retained per viewport key, not unbounded history, not collaborative CRDT merge. |
| 5 | Formula language-service primitives | Needs one more fold-in for safe code actions | Formula assist, spans, diagnostics, hover, completions, and reference cycling are valuable, but the strong claim is not safe rename yet. | Fixture: LET names, sheet-qualified refs, structured refs, external refs, dynamic arrays, shared formulas, 3D refs. Benchmark: assist latency on large formula corpora. Surface: CLI/API plus SDK; MCP if agents need formula repair. Validation gate: parser span tests, assist tests, lint/trace integration, oracle cases. Competitor contrast: HyperFormula has a mature parser, AST, dependency graph, and hundreds of functions; Ascend should claim explainable edit primitives over workbook-preserving files, not function-count parity. Honest boundary: rename/code actions are previews until binding resolution proves safety. |
| 6 | Release proof bundle | Needs one more fold-in, piggybacks rank 2 | This is the product wrapper for safe opening and auditable mutation. It is high value but should not be built before the package-part proof schema is stable. | Fixture: one public real workbook per scenario with inspect, plan, commit, reopen, diff, audit, hashes. Benchmark: bundle generation overhead. Surface: CLI/API artifact and SDK schema. Validation gate: fixture-backed golden proof bundle tests. Competitor contrast: generic libraries can read/write files; Ascend can show the whole decision trail. Honest boundary: no signed provenance unless signatures and verifier roots exist. |
| 7 | Formula conformance/oracle routing | One correctness loop away | The route fields exist, but the product claim is internal until there are runnable oracle adapters and completed artifacts. | Fixture: static formula corpus plus Excel/LibreOffice/HyperFormula routed mismatches. Benchmark: oracle routing overhead and corpus completion time. Surface: JSON corpus output first; later CLI report. Validation gate: mismatch-class tests and completed artifact verifier. Competitor contrast: HyperFormula is a direct formula engine baseline; Ascend should prove where it agrees, diverges, or delegates. Honest boundary: cached workbook values are not always truth. |
| 8 | Columnar scan sidecars | Speculative for product, useful for performance research | The benchmark harness is useful, but the implementation is not yet part of production workbook semantics. | Fixture: real tables and ranges with mixed types, dates, blanks, formulas, filters. Benchmark: repeated scans, sidecar build cost, invalidation cost, memory. Surface: benchmark harness first, maybe SDK internal cache later. Validation gate: checksum parity and generation invalidation tests. Competitor contrast: DuckDB already reads XLSX into typed table-shaped scans; Ascend’s angle is disposable sidecars over a preservation-first workbook model. Honest boundary: not a replacement storage engine and not a universal workbook rewrite. |

## Claim Status Buckets

### Credibly Claim Today

- Safe unknown workbook opening, if phrased as "Ascend recommends a load mode and trust review before hydrating unknown XLSX/XLSM files."
- Token-bounded agent view, if phrased as "Ascend can emit deterministic, budgeted workbook summaries with omission counters."
- Retained viewport patch history for SDK/compact-read surfaces, if phrased with bounded retention and explicit invalidation reasons.

### Needs One More Fold-In

- Auditable package-part mutation: needs a stable proof bundle over passthrough/regenerate/add/drop/error with real fixtures.
- Release proof bundle: needs the package-part proof schema and fixture-backed CLI/API artifact.
- Formula language-service primitives: needs parser-native binding roles before safe rename/code-action claims.
- Formula conformance/oracle routing: needs completed artifacts and runnable oracle adapters.

### Still Speculative

- Columnar scan sidecars as a product claim.
- Collaborative or CRDT-like spreadsheet editing.
- Signed release provenance or third-party attestation.
- "Excel-compatible formula engine" as a blanket claim.

## Top Handoffs Only

### 1. Correctness/Product: Auditable Package-Part Mutation

```text
/goal Promote auditable package-part mutation as the next correctness/product proof. Do not add broad writer behavior. Define a stable PackagePartProof schema over existing plan/commit evidence with actions passthrough, regenerate, add, drop, and error. Prove it on public fixtures covering docProps passthrough, generated worksheet XML, calc-chain drop, signature invalidation, and macro/ActiveX preservation blocks. Expose compact proof summaries in CLI/API/MCP and full SDK JSON. Validate with targeted tests, bunx tsc --build, bunx biome check, and bun run test:changed.
```

### 2. Product: Safe Unknown Workbook Opening Proof Bundle

```text
/goal Turn safe unknown workbook opening into a product proof bundle. Use existing open-plan SDK/CLI/API/MCP surfaces; do not add another narrow surface. Build a fixture-backed report that shows package fingerprint, recommended load mode, reviewBeforeHydration reasons, trust/active-content next step, and measured latency versus full hydration on clean, macro, signature, pivot/slicer, and malformed workbooks. Keep boundaries explicit: this is pre-hydration risk routing, not malware scanning. Validate docs, focused tests, typecheck, biome, and changed tests.
```

## Next-Loop Prompts

### Correctness Loop

```text
/goal Build the auditable package-part mutation proof. Start from current package action proof surfaces and produce one stable schema plus fixture-backed tests. The deliverable is proof evidence for passthrough/regenerate/add/drop/error and honest loss boundaries, not a writer rewrite.
```

### Performance Loop

```text
/goal Benchmark safe unknown workbook opening as a product claim. Measure open-plan package fingerprint latency against full hydration across clean, macro, pivot/slicer, and malformed public fixtures. Add no production behavior unless the benchmark exposes a tiny instrumentation gap with an obvious owner.
```

### Product Loop

```text
/goal Package token-bounded agent view and retained viewport patch history as documented agent contracts. Focus on examples, omission/invalidation metadata, and recovery guidance across existing SDK/CLI/API/MCP surfaces. Do not expand retention or add new protocols until telemetry proves the need.
```

## Do Not Promote Yet

- Columnar scan sidecars: keep in the performance loop until real workbook table benchmarks prove speedups after build and invalidation costs.
- Safe formula rename: keep as research until AST spans, symbol roles, and binding resolution prove no accidental edits to strings, external refs, or structured refs.
- Release proof bundle with provenance language: do not imply signed, tamper-evident, or SLSA-style guarantees.
- Universal formula compatibility: use routed oracle evidence, not blanket Excel parity.
- Collaborative spreadsheet sync: retained viewport history is not CRDT collaboration.
- Private workbook corpus results: never use private or large workbook data as public proof.
