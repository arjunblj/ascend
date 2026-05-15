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

Rank order is intentional. The first two claims are the only implementation-loop handoffs from this synthesis block. Every row is product-shaped, not surface-shaped.

| Claim wording allowed today | Proof still missing | Owner loop |
| --- | --- | --- |
| **Safe unknown workbook opening:** Ascend can inspect XLSX package features before full workbook hydration and recommend a load mode or review step across SDK, CLI, API, and MCP. | **Fixture:** release public clean, formula-heavy, XLSM macro, pivot/slicer, ActiveX/chart, signed, unknown-part, and malformed workbooks; replace synthetic signed/unknown edge cases with durable public binaries if available.<br>**Benchmark:** release-environment open-plan package fingerprint latency versus full hydration with package size, part count, and chosen mode.<br>**Surface:** existing SDK open planner, CLI open-plan, API endpoint, and MCP tool only; no new surface.<br>**Validation gate:** focused safe-open tests, report generation, malformed package checks, typecheck/Biome if code changes.<br>**Competitor contrast:** Microsoft Protected View is trust UX; Ascend's claim is OSS pre-hydration package-feature routing.<br>**Honest boundary:** not malware scanning, sandboxing, file trust, or proof that active content is safe. | **Product/performance. Top handoff.** |
| **Auditable package-part mutation:** Ascend can explain write outcomes with structured package action evidence and rollback-journal issue reasons using `passthrough`, `regenerate`, `add`, `drop`, and `error`. | **Fixture:** release public docProps passthrough, worksheet regeneration, calc-chain drop, signature invalidation, macro/ActiveX preservation, drawing/chart sidecar accounting, and unknown-part rejection.<br>**Benchmark:** package-proof overhead in bytes and milliseconds for plan/commit with compact versus expanded evidence.<br>**Surface:** full SDK evidence plus existing compact CLI/API/MCP proof summaries and opt-in expansion; no new mutation surface.<br>**Validation gate:** package-action harness, plan/commit/reopen/diff/audit tests, journal/package compatibility, schema snapshot, typecheck/Biome for code changes.<br>**Competitor contrast:** openpyxl and SheetJS document preservation boundaries; Ascend's stronger claim is per-part accounting.<br>**Honest boundary:** not signed provenance, tamper-evident attestation, SLSA, or Excel semantic refresh proof; chart XML is regenerated while drawing sidecars pass through, so do not claim chart byte passthrough. | **Correctness/product. Top handoff.** |
| **Token-bounded agent view:** Ascend emits deterministic workbook summaries with requested budget metadata, estimated token counts, unbudgeted estimate, omission counters, compact omitted-evidence locators, and formula-pattern example refs. | **Fixture:** current proof covers dense table, wide sparse, formula-heavy, metadata-heavy, and public formula-stress workbooks.<br>**Benchmark:** tracked harness reports full versus budgeted estimates, compression ratios, omitted rows/values/formulas, and shape preservation.<br>**Surface:** existing SDK, CLI, API, and MCP agent-view/read surfaces only.<br>**Validation gate:** deterministic truncation, cross-surface JSON shape, omitted-evidence locator recovery, and no hidden summarization without counters are covered in the current rerun.<br>**Competitor contrast:** Univer exposes agent spreadsheet operations; Ascend's claim is deterministic local evidence under token budgets.<br>**Honest boundary:** token counts are approximate; omitted evidence is absent by design; agent view does not replace package inspection or proof artifacts; wide sparse ranges can exceed tiny requested budgets because column summaries are the structural floor. | Product/DX. Product-proof backed; release copy still needs a concrete example. |
| **Retained viewport patch history:** Ascend can patch SDK interactive viewports from bounded retained tokens and return explicit invalidation reasons when it cannot; API and MCP compact reads expose `changedSince` invalidation recovery. | **Fixture:** unchanged viewport, changed cells/styles/metadata, invalid token, cross-session token, expired token, skipped token, changed projection, and source invalidation.<br>**Benchmark:** patch bytes and latency versus full viewport refresh under repeated edits and metadata invalidation.<br>**Surface:** SDK interactive patch stream plus API/MCP compact recovery only; CLI is explicitly excluded today.<br>**Validation gate:** viewport proof harness, SDK interactive contract tests, API/MCP compact `changedSince` tests, retention cap assertions.<br>**Competitor contrast:** database MVCC retains readable versions; this is not a CRDT or collaborative editing engine.<br>**Honest boundary:** bounded per-window history only, not unlimited history, multi-writer sync, or transaction isolation across all workbook metadata. | Product/performance. |
| **Formula language-service primitives:** Ascend exposes parse diagnostics, token/reference spans, hover, completions, reference cycling, binding roles, and a guard that refuses unsafe rename targets. | **Fixture:** LET declarations/uses with shadowing, workbook/sheet defined names, table names, table columns, structured item selectors, external workbook refs, 3D refs, spill refs, dynamic arrays, shared formulas, literals, punctuation, and parse failures.<br>**Benchmark:** assist latency on long formulas and formula corpora.<br>**Surface:** existing SDK `formulaAssist`, CLI formula assist, API formula-assist, and MCP formula-assist only.<br>**Validation gate:** parser/span/binding-role tests, cross-surface assist tests, and rejection matrix tests. Current `packages/sdk/src/formula-edit.test.ts` passes.<br>**Competitor contrast:** HyperFormula owns formula engine/dependency graph breadth; Ascend should claim workbook-preserving edit intelligence.<br>**Honest boundary:** **no edit-producing rename** and no safe workbook-context rename claim. | Product/DX plus correctness. Rejection-only in this block. |
| **Release proof bundle:** Ascend has ingredients for inspect, plan, commit, reopen, diff, audit, digests, package action proof, and explicit boundaries. | **Fixture:** one real public workbook workflow per top claim.<br>**Benchmark:** bundle size and generation overhead compared with normal commit/report flow.<br>**Surface:** stable SDK schema first, report generation second; CLI/API/MCP references only after artifact storage and privacy semantics stabilize.<br>**Validation gate:** golden proof fixtures, digest checks, reopen/diff/audit checks, package graph audit checks, and failure cases.<br>**Competitor contrast:** generic libraries read/write files; Ascend explains the decision trail.<br>**Honest boundary:** not signed, tamper-evident, SLSA, in-toto, certified provenance, or third-party attestation. | Product after ranks 1 and 2 stabilize. |
| **Formula conformance/oracle routing:** Ascend can classify formula mismatch classes in research and route next oracle work. | **Fixture:** completed corpus by mismatch class: cached-only, volatile, numeric drift, unsupported function, external refs, dynamic arrays, structured refs, and date-system behavior.<br>**Benchmark:** corpus completion time and per-oracle route overhead.<br>**Surface:** completed JSON artifacts and CLI report only; no MCP/API promotion yet.<br>**Validation gate:** converter tests, artifact verifier, skipped/divergence counters, and no threshold changes without evidence.<br>**Competitor contrast:** HyperFormula is the strongest OSS formula baseline; Excel/LibreOffice are behavior oracles with automation limits.<br>**Honest boundary:** no blanket Excel-compatible formula claim. | Correctness backlog. |
| **Columnar scan sidecars:** Ascend has research evidence that disposable sidecars may accelerate repeated scans without replacing workbook truth. | **Fixture:** real workbook tables/ranges with numbers, dates, blanks, strings, formulas, filters, hidden rows, and styles.<br>**Benchmark:** repeated scans, sidecar build cost, invalidation cost, memory overhead, and checksum parity against canonical workbook reads.<br>**Surface:** benchmark harness only; no SDK/CLI/API/MCP product surface.<br>**Validation gate:** generation-key invalidation, checksum parity, memory cap tests, and benchmark guard before production.<br>**Competitor contrast:** DuckDB reads XLSX ranges into typed SQL tables; Arrow supplies the columnar scan substrate.<br>**Honest boundary:** not a storage engine, not a workbook rewrite, and not guaranteed faster for sparse or single-pass reads. | Performance research only. Do not promote. |

## Claim Ladder

### Credible Today

1. **Safe unknown workbook opening** can be claimed in guarded language because the feature-fingerprint/open-plan path exists across surfaces and has a tracked proof harness. The next loop should package evidence, not invent another opener.
2. **Auditable package-part mutation** can be claimed in guarded language because package action evidence and journal issue reasons exist. The next loop should make the proof report durable and compact.
3. **Token-bounded agent view** can be claimed for deterministic summaries, omission metadata, locator recovery, and formula-pattern example recovery. It still needs a concise product example before release headline copy.

### Needs One More Fold-In

4. **Retained viewport patch history** needs one proof pass that ties SDK retention semantics to any public API/CLI/MCP surface before release copy.
5. **Formula language-service primitives** need workbook-context refusal proof across API/CLI/MCP and a latency/corpus gate. Do not add rename; strengthen rejection evidence only.
6. **Release proof bundle** needs artifact storage, privacy boundaries, and stable report generation before it is more than a benchmark index.

### Still Speculative

7. **Formula conformance/oracle routing** is a correctness research program until mismatch classes are complete and reproducible without private corpora.
8. **Columnar scan sidecars** are performance research until invalidation, memory caps, and parity are proven on real workbook-shaped ranges.

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

## Next-Loop Prompts

### Product/Performance Loop

```text
/goal Become Ascend's safe-open release-proof owner. Do not add new open surfaces. Use the existing feature-fingerprint/open-plan implementation and proof harness to produce a public-fixture evidence report for "safe unknown workbook opening." Required proof: fixture list, latency versus full hydration, SDK/CLI/API/MCP surface evidence, validation command, competitor contrast with Microsoft Protected View, and boundaries saying this is not malware scanning, sandboxing, or active-content safety. Commit only the report, harness fixes if needed, and index updates.
```

### Correctness/Product Loop

```text
/goal Become Ascend's auditable package-part mutation proof owner. Do not add new mutation surfaces. Use existing package action evidence, release proof index, and rollback journal issue reasons to produce a compact report for "auditable package-part mutation." Required proof: per-part passthrough/regenerate/add/drop/error fixtures, journal/package compatibility, SDK evidence shape, validation gate, competitor contrast with openpyxl/SheetJS preservation boundaries, and honest boundaries around signatures, chart byte passthrough, and provenance. Commit only proof/report/harness fixes and index updates.
```

### Product/DX Loop

```text
/goal Become Ascend's formula intelligence claim steward. Do not implement rename. Strengthen only rejection-first evidence for formula language-service primitives: binding roles, LET shadowing, defined names, table names, table columns, external refs, 3D refs, spills, parse failures, and cross-surface refusal reasons. The output is a proof report and tests for refusal semantics, not an edit-producing code action.
```

## Do Not Promote Yet

- Formula safe rename or edit-producing rename.
- Columnar sidecars as a product feature.
- Release proof bundle as signed provenance.
- Universal Excel formula compatibility.
- Collaborative/CRDT claims from retained viewport patches.
- Claims backed only by private workbook corpora.
