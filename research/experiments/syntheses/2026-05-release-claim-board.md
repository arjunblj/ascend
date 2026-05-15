# 2026-05 Release Claim Board

Date: 2026-05-15

## Purpose

Freeze the next release claims into proof-shaped language. This board is deliberately conservative: it states what Ascend may say today, what proof is still missing, and which owner loop should carry the next step. Research should not keep promoting new production surfaces while these proof gaps remain open.

## External References

- Microsoft Protected View frames unsafe file opening as read-only review with active-content restrictions: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions define packages as parts plus relationships and describe digital signatures over package contents: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl documents preservation boundaries for unsupported workbook objects: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS CE write docs describe writer scope and data-preservation orientation: https://docs.sheetjs.com/docs/api/write-options/
- LSP 3.17 separates `prepareRename` from edit-producing rename and allows a server to refuse by returning no target: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft names in formulas document workbook- and worksheet-scoped names: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft structured references document table and column symbols in formulas: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- HyperFormula documents formula references, dependency graphs, and formula-engine scope: https://hyperformula.handsontable.com/guide/cell-references.html
- Apache Arrow describes columnar layout as an analytics-oriented memory format: https://arrow.apache.org/docs/format/Columnar.html
- DuckDB documents direct XLSX range/table ingestion through `read_xlsx`: https://duckdb.org/docs/stable/guides/file_formats/excel_import

## Release Claim Board

| Claim wording allowed today | Proof still missing | Owner loop |
| --- | --- | --- |
| **Safe unknown workbook opening:** Ascend can inspect workbook package features before full hydration and recommend a load mode or review step from SDK, CLI, API, and MCP surfaces. | Current checkpoint: tracked safe-open proof harness covers public clean, formula-heavy, XLSM macro, pivot, ActiveX, chart, synthetic signed, synthetic unknown-part, and malformed cases with latency versus full hydration. Missing before stronger release copy: release-environment rerun, durable public binary signed/unknown fixtures if available, and one published report generated from existing surfaces. Competitor contrast: Protected View is trust UX; Ascend's claim is OSS pre-hydration package-feature routing. Boundary: not malware scanning, sandboxing, or active-content safety. | **Product/performance.** Top handoff. |
| **Auditable package-part mutation:** Ascend can explain write outcomes with structured package action evidence and rollback-journal issue reasons. | Current checkpoint: commit-local proofs infer source graph evidence from retained source bytes, and the tracked proof harness covers `passthrough`, `regenerate`, `add`, `drop`, and `error` over docProps, generated sheet XML, calc-chain, signatures, macros, chart/drawing accounting, and unknown-part rejection. Missing before stronger release copy: release-environment rerun, durable public binary fixtures for synthetic edge packages if available, and a compact published report that preserves full SDK evidence. Competitor contrast: openpyxl and SheetJS expose preservation boundaries; Ascend's stronger claim is per-part accounting. Boundary: not signed provenance or Excel semantic refresh proof; chart XML is regenerated while drawing sidecars pass through, so do not claim chart byte passthrough. | **Correctness/product.** Top handoff. |
| **Token-bounded agent view:** Ascend emits deterministic workbook summaries with requested budget metadata, estimated token counts, unbudgeted estimate, and omission counters. | Fixture: wide sparse, dense table, formula-heavy, metadata-heavy, and mixed-sheet public workbooks. Benchmark: raw JSON versus full agent view versus budgeted agent view. API/CLI/MCP: existing agent-view surfaces only. Validation: deterministic truncation, cross-surface JSON shape, no omitted evidence without counters. Current checkpoint: tracked harness proves deterministic shape preservation and omission counters; very wide sparse ranges can exceed tiny requested budgets because column summaries are preserved as the structural floor. Competitor contrast: Univer exposes agent spreadsheet operations; Ascend's claim is deterministic local evidence under token budgets. Boundary: token counts are approximate; omitted evidence is absent by design; agent view does not replace metadata inspection. | Product/DX. Hold behind top two. |
| **Retained viewport patch history:** Ascend can patch from bounded retained tokens and return explicit invalidation reasons when it cannot. | Fixture: changed cells/styles/metadata, stale token, invalid token, expired token, skipped token, projection change, cross-session token. Benchmark: patch bytes and latency versus full refresh. API/CLI/MCP: prove how existing compact/interactive APIs expose tokens before marketing this broadly. Validation: compact changedSince tests, interactive viewport tests, retention cap tests. Current checkpoint: tracked SDK interactive harness covers retained patch, skipped token, invalid/cross-session/expired/projection/metadata invalidations; projection change currently returns `base-snapshot-missing`, not a dedicated projection code. Competitor contrast: MVCC retains readable versions; most spreadsheet OSS exposes snapshots. Boundary: bounded patch retention, not CRDT collaboration or unlimited history. | Product/performance. |
| **Formula language-service primitives:** Ascend exposes parse diagnostics, token/reference spans, hover, completions, reference cycling, binding roles, and a guard that refuses unsafe rename targets. | Fixture: LET declarations/uses with shadowing, workbook/sheet defined names, table names, table columns, external workbook refs, 3D refs, dynamic arrays, shared formulas, parse failures. Benchmark: assist latency on long formulas and formula corpora. API/CLI/MCP: current formula-assist surfaces only. Validation: parser/span/binding-role tests and cross-surface assist tests. Current checkpoint: SDK tests prove lexical LET shadowing, formula-local prepare evidence, and refusal for unresolved/table/cell references; no edit-producing rename exists. Competitor contrast: HyperFormula owns engine/dependency graph breadth; Ascend should claim workbook-preserving edit intelligence. Boundary: **no edit-producing rename** and no safe workbook-context rename claim. | Product/DX plus correctness, but no rename implementation in this block. |
| **Release proof bundle:** Ascend has ingredients for inspect, plan, commit, reopen, diff, audit, digests, package action proof, and boundaries. | Fixture: one real public workbook workflow per top claim. Benchmark: bundle size and generation overhead. API/CLI/MCP: stable SDK schema first, report generation second. Validation: golden proof bundle fixtures, digest checks, reopen/diff/audit checks, explicit failure cases. Current checkpoint: safe-open and package-action harnesses should stay as sibling suite-level release evidence artifacts, referenced by digest later rather than embedded into each workbook-level proof. Competitor contrast: generic libraries read/write; Ascend explains the decision trail. Boundary: not signed, tamper-evident, SLSA, or certified provenance. | Product after safe-open and package-part schemas stabilize. |
| **Formula conformance/oracle routing:** Ascend can classify formula mismatch classes in research and route next oracle work. | Fixture: completed corpus by mismatch class: cached-only, volatile, numeric drift, unsupported function, external refs, dynamic arrays, structured refs, date system. Benchmark: corpus completion and route overhead. API/CLI/MCP: none yet; CLI report is enough when ready. Validation: converter tests, artifact verifier, skipped/divergence counters. Competitor contrast: HyperFormula is the strongest OSS formula baseline; Excel/LibreOffice are behavior oracles with automation limits. Boundary: no blanket Excel-compatible formula claim. | Correctness backlog. |
| **Columnar scan sidecars:** Ascend has research evidence that disposable sidecars may accelerate repeated scans without replacing workbook truth. | Fixture: real workbook tables/ranges with numbers, dates, blanks, strings, formulas, filters, hidden rows, and styles. Benchmark: repeated scans, build cost, invalidation cost, memory, checksum parity. API/CLI/MCP: benchmark harness only. Validation: generation-key invalidation, memory cap, parity. Competitor contrast: DuckDB reads XLSX into typed SQL; Arrow supplies columnar layout. Boundary: not a storage engine or faster single-pass guarantee. | Performance research only. Do not promote. |

## Formula Intelligence Rejection-First Spec

This is a spec for refusal, not an implementation request. Do not add edit-producing rename in this block.

### Binding Roles

Formula intelligence may classify stateless formula-local symbols into:

- `let-binding-declaration`: the declaration name in a `LET` binding pair.
- `let-binding-use`: a resolved use of the nearest visible `LET` declaration.
- `table-name-use`: the table identifier portion of a structured reference.
- `table-column-use`: the column identifier portion of a structured reference.
- `unresolved-name`: a name-like token that needs workbook context before it can be identified.

Every role must preserve the original formula span. A role is not enough to authorize mutation unless the target is formula-local and all affected occurrences are known.

### LET Shadowing

`LET` binding resolution is lexical and nearest-scope-wins. Nested `LET` declarations with the same text shadow outer declarations only inside the nested body. A guard may prepare a local target only when it can prove the declaration span and every use span for that exact binding. It must not rewrite other equal text in outer or inner scopes.

Allowed guard result today:

- Cursor on a resolved `LET` declaration or use.
- Result names the declaration range, placeholder, and all formula-local occurrence ranges.
- Result does not apply edits.

Required refusal:

- Cursor on a `LET` token, function name, separator, whitespace, literal, or operator.
- Cursor on a binding use whose declaration span cannot be proven.
- Cursor on an equal-text name outside the proven lexical binding.

### Defined Names

Defined names require workbook context because Excel supports workbook-scoped and worksheet-scoped names, plus sheet-qualified names. A stateless formula token such as `Budget` cannot prove whether the target is a workbook name, sheet-local name, table name, function, or missing name.

`prepareRename` must refuse defined names unless a future workbook-context resolver proves:

- exact scope: workbook versus worksheet;
- visible sheet context for formula location;
- all formula, chart, validation, conditional-format, table, and defined-name references that would need edits;
- collision rules and Excel name validity;
- external/workbook-index qualification is absent or intentionally unsupported.

Current reason: `workbook-context-required`.

### Table Names And Table Columns

Structured references contain table and column symbols with escaping and special item syntax. Excel updates structured references when tables or columns are renamed, which makes a safe rename workbook-wide, not formula-local.

`prepareRename` must refuse:

- table names such as `Sales` in `Sales[Amount]`;
- table columns such as `Amount` in `Sales[Amount]`;
- current-row and item selectors such as `[@[Amount]]`, `[#Totals]`, and nested structured-reference segments;
- ambiguous escaped columns, duplicate/invalid table metadata, or query-backed columns.

Current reason: `workbook-context-required`.

### External, Sheet, Cell, And Range References

Cell, range, sheet-qualified, 3D, spill, and external workbook references are workbook/path operations, not symbol rename targets. They may be cycled, highlighted, linted, or rewritten by explicit workbook operations, but not prepared as local rename.

`prepareRename` must refuse:

- `A1`, `$A$1`, `A1:B2`, whole-row/whole-column references;
- `Sheet1!A1`, `'My Sheet'!A1`, `Sheet1:Sheet3!A1`;
- `[Book.xlsx]Sheet1!A1`, workbook-index references, or quoted external paths;
- spill references such as `A1#`;
- references inside validations, conditional formats, chart series, defined names, and table formulas unless a future operation-specific planner owns the rewrite.

Current reason: `reference-target-not-renameable`.

### Prepare-Rename Contract

The guard may return `ok: true` only for formula-local `LET` bindings with complete lexical evidence. It must return `ok: false` for anything requiring workbook context or reference semantics.

| Cursor target | Required result | Reason |
| --- | --- | --- |
| Resolved `LET` declaration/use | `ok: true`, declaration range, placeholder, occurrence ranges | Formula-local only; no edit application. |
| Workbook or sheet defined name | `ok: false` | `workbook-context-required` |
| Table name or table column | `ok: false` | `workbook-context-required` |
| Cell/range/sheet/3D/spill/external ref | `ok: false` | `reference-target-not-renameable` |
| Function name or formula keyword | `ok: false` | `no-symbol-at-cursor` or future function-specific refusal |
| Literal, operator, punctuation, whitespace | `ok: false` | `no-symbol-at-cursor` |
| Parse failure or ambiguous tokenization | `ok: false` | `no-symbol-at-cursor` unless a diagnostic span is separately reported |

## Top Handoffs

1. Product/performance: prove safe unknown workbook opening from existing open-plan surfaces over public fixtures and latency evidence.
2. Correctness/product: prove auditable package-part mutation with stable per-part outcomes and journal/package proof compatibility.

## Do Not Promote Yet

- Formula safe rename or edit-producing rename.
- Columnar sidecars as a product feature.
- Release proof bundle as signed provenance.
- Universal Excel formula compatibility.
- Collaborative/CRDT claims from retained viewport patches.
- Claims backed only by private workbook corpora.
