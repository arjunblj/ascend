# 2026-05 Release Claim Board

Date: 2026-05-15

## Purpose

Freeze the next release claims into proof-shaped language. This board is deliberately conservative: it states what Ascend may say today, what proof is still missing, and which owner loop should carry the next step. Research should not keep promoting new production surfaces while these proof gaps remain open.

Current stewardship rule: hand off only the top one or two product claims to implementation loops. Everything else stays as proof packaging, validation, or "do not promote yet" until the missing evidence is explicit.

## External References

- Microsoft Protected View frames unsafe file opening as read-only review with active-content restrictions: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions define packages as parts plus relationships and describe digital signatures over package contents: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl documents preservation boundaries for unsupported workbook objects: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS CE write docs describe writer scope and data-preservation orientation: https://docs.sheetjs.com/docs/api/write-options/
- LSP 3.17 separates `prepareRename` from edit-producing rename and allows a server to refuse by returning no target: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft LET documents formula-local names scoped to the LET function: https://support.microsoft.com/en-gb/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
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
| **Retained viewport patch history:** Ascend can patch SDK interactive viewports from bounded retained tokens and return explicit invalidation reasons when it cannot; API and MCP compact reads expose `changedSince` invalidation recovery. | **Fixture:** current proof covers retained patch, skipped token, invalid token, cross-session token, expired history, projection change, metadata invalidation, changed window, selected-column projection, and changed source.<br>**Benchmark:** tracked harness reports patch bytes; retained patch cases total 630 bytes in the latest rerun.<br>**Surface:** SDK interactive patch stream plus API/MCP compact recovery only; CLI is explicitly excluded today.<br>**Validation gate:** viewport proof harness, SDK interactive contract tests, API/MCP compact `changedSince` tests, retention cap assertions.<br>**Competitor contrast:** database MVCC retains readable versions; this is not a CRDT or collaborative editing engine.<br>**Honest boundary:** bounded per-window history only, not unlimited history, multi-writer sync, or transaction isolation across all workbook metadata. | Product/performance. Product-proof backed with CLI excluded. |
| **Formula language-service primitives:** Ascend exposes formula IDE building blocks: parse diagnostics, token/reference spans, hover, completions, reference cycling, binding roles, and a corpus-backed rejection-first prepare-rename guard. Allowed wording must say the guard refuses workbook-context and reference targets; it must not imply edit-producing rename. | **Fixture:** SDK rejection matrix covers LET declarations/uses with shadowing, workbook-context names, table names, table columns, structured item selectors, external workbook refs, 3D refs, spill refs, literals, punctuation, and parse failures. The proof harness samples public POI/ClosedXML formula workbooks.<br>**Benchmark:** `fixtures/benchmarks/formula-assist-proof.ts` currently discovers 1685 public formulas, samples 1685, reports 2322 reference spans, 25 binding roles, 3 LET-local prepare-rename OK targets, and 1692 prepare-rename refusals by reason. Latest local P95 assist latency was 0.0368 ms on this machine.<br>**Surface:** existing SDK `formulaAssist`, CLI formula assist, API formula-assist, and MCP formula-assist only; no new rename surface.<br>**Validation gate:** parser/span/binding-role tests, cross-surface assist tests, and the formula-assist proof harness. Latest local proof passed; no production code was changed in this stewardship refresh.<br>**Competitor contrast:** HyperFormula owns formula engine/dependency graph breadth; Ascend should claim workbook-preserving edit intelligence and refusal semantics, not a broader formula IDE.<br>**Honest boundary:** **no edit-producing rename**, no safe workbook-context rename, no table/defined-name rename from formula assist, and no claim that all formula references can be rewritten safely. | Product/DX stewardship only. Proof/spec boundary; not an implementation handoff. |
| **Release proof bundle:** Ascend has ingredients for inspect, plan, commit, reopen, diff, audit, digests, package action proof, and explicit boundaries. | **Fixture:** one real public workbook workflow per top claim.<br>**Benchmark:** bundle size and generation overhead compared with normal commit/report flow.<br>**Surface:** stable SDK schema first, report generation second; CLI/API/MCP references only after artifact storage and privacy semantics stabilize.<br>**Validation gate:** golden proof fixtures, digest checks, reopen/diff/audit checks, package graph audit checks, and failure cases.<br>**Competitor contrast:** generic libraries read/write files; Ascend explains the decision trail.<br>**Honest boundary:** not signed, tamper-evident, SLSA, in-toto, certified provenance, or third-party attestation. | Product after ranks 1 and 2 stabilize. |
| **Formula conformance/oracle routing:** Ascend can classify formula mismatch classes in research and route next oracle work. | **Fixture:** completed corpus by mismatch class: cached-only, volatile, numeric drift, unsupported function, external refs, dynamic arrays, structured refs, and date-system behavior.<br>**Benchmark:** corpus completion time and per-oracle route overhead.<br>**Surface:** completed JSON artifacts and CLI report only; no MCP/API promotion yet.<br>**Validation gate:** converter tests, artifact verifier, skipped/divergence counters, and no threshold changes without evidence.<br>**Competitor contrast:** HyperFormula is the strongest OSS formula baseline; Excel/LibreOffice are behavior oracles with automation limits.<br>**Honest boundary:** no blanket Excel-compatible formula claim. | Correctness backlog. |
| **Columnar scan sidecars:** Ascend has research evidence that disposable sidecars may accelerate repeated scans without replacing workbook truth. | **Fixture:** real workbook tables/ranges with numbers, dates, blanks, strings, formulas, filters, hidden rows, and styles.<br>**Benchmark:** repeated scans, sidecar build cost, invalidation cost, memory overhead, and checksum parity against canonical workbook reads.<br>**Surface:** benchmark harness only; no SDK/CLI/API/MCP product surface.<br>**Validation gate:** generation-key invalidation, checksum parity, memory cap tests, and benchmark guard before production.<br>**Competitor contrast:** DuckDB reads XLSX ranges into typed SQL tables; Arrow supplies the columnar scan substrate.<br>**Honest boundary:** not a storage engine, not a workbook rewrite, and not guaranteed faster for sparse or single-pass reads. | Performance research only. Do not promote. |

## Claim Ladder

### Credible Today

1. **Safe unknown workbook opening** can be claimed in guarded language because the feature-fingerprint/open-plan path exists across surfaces and has a tracked proof harness. The next loop should package evidence, not invent another opener.
2. **Auditable package-part mutation** can be claimed in guarded language because package action evidence and journal issue reasons exist. The next loop should make the proof report durable and compact.
3. **Token-bounded agent view** can be claimed for deterministic summaries, omission metadata, locator recovery, and formula-pattern example recovery. It still needs a concise product example before release headline copy.
4. **Retained viewport patch history** can be claimed for SDK retained patches plus API/MCP compact recovery, with CLI explicitly excluded.

### Needs One More Proof Package

5. **Formula language-service primitives** are allowed only as a corpus-backed, rejection-first primitives claim. Cross-surface refusal snapshots and a latency/corpus proof now exist. Do not add rename.
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
/goal Become Ascend's safe-open release-proof owner. Do not add new open surfaces. Start from `fixtures/benchmarks/release-proof-index.ts` and the `safe-open-proof` artifact. Reproduce with `bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json` and a timed local rerun when needed. Required proof: fixture list, package counts, recommended load mode, reviewBeforeHydration, risk families, malformed rejection, SDK/CLI/API/MCP surface evidence, validation commands, stable shape digest, competitor contrast with Microsoft Protected View, and boundaries saying this is not malware scanning, sandboxing, file trust, active-content safety, signed provenance, or a release performance threshold. Resolve or explicitly accept the current blockers: signed/unknown-part cases are code-generated packages rather than public binary fixtures, and timing evidence is local proof-run data. Commit only release report/index updates or narrow harness fixes.
```

### Correctness/Product Loop

```text
/goal Become Ascend's auditable package-part mutation proof owner. Do not add new mutation surfaces. Start from `fixtures/benchmarks/release-proof-index.ts` and the `package-action-proof` artifact. Reproduce with `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json` and a timed local rerun when needed. Required proof: per-part `passthrough`/`regenerate`/`add`/`drop`/`error` coverage, source graph evidence, journal package issue refs, SDK evidence shape, validation commands, stable shape digest, competitor contrast with OPC/openpyxl/SheetJS, and honest boundaries around signatures, chart byte passthrough, Excel recalculation equivalence, SLSA, in-toto, and provenance. Resolve or explicitly accept the current blockers: synthetic edge packages must stay disclosed unless replaced by public binary fixtures, and the proof is local evidence rather than signed attestation. Commit only proof/report/harness fixes and index updates.
```

### Proof-Owner Exit Criteria

The owner loops above are done only when the release artifact can answer these checks without private data:

| Owner | Required evidence | Blocking exit condition |
| --- | --- | --- |
| Safe unknown workbook opening | `safe-open-proof` command, stable shape digest `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`, public fixture table, explicit review branches, malformed rejection, validation commands, and accepted or resolved `readyWhen` gates | Do not publish headline copy if signed/unknown package evidence remains undisclosed or if timing language reads as a threshold |
| Auditable package-part mutation | `package-action-proof` command, stable shape digest `9abebf576651551f58e00ccf8469d099b2c06dacd48391fe581a24e51a1e0afd`, all five action classes, one representative streaming writer proof, journal issue refs, post-write audit status, validation commands, and accepted or resolved `readyWhen` gates | Do not publish headline copy if synthetic edge packages are hidden, if chart XML is called byte-passthrough, if one streaming proof is described as full streaming matrix coverage, or if the proof implies SLSA/in-toto/signed provenance |

### Public Fixture Policy

Generated edge packages are acceptable release proof only when all of these are true:

- The edge case is structural and package-level, not dependent on private workbook content.
- The generated package is built by tracked code in a proof harness.
- The report labels the case as synthetic or generated.
- The release proof index carries a publication blocker until product explicitly accepts generated proof.
- A fixture scan has checked the public corpus for a binary replacement.

Public binary fixtures are required before stronger headline copy when any of these are true:

- The claim depends on real-world workbook authoring behavior rather than package topology.
- The edge case involves vendor-specific semantics, UI behavior, or Excel repair behavior.
- The proof would otherwise imply trust, malware scanning, signed provenance, or third-party attestation.
- The generated fixture would hide licensing, privacy, or provenance uncertainty.

Current application: the safe-open signed and unknown-part cases can remain generated local proof because they exercise OPC package topology and are disclosed as synthetic. They should still block stronger release copy until accepted or replaced; `fixtures/benchmarks/safe-open-fixture-scan.ts` currently finds no checked-in public binary replacement.

Release-index enforcement: `fixtures/benchmarks/release-proof-index.ts` marks both current top artifacts as `headlineClaimAllowed: false` with `releaseGate: blocked-by-publication-policy`. It exposes per-artifact `readyWhen` requirements by owner loop and an aggregate `readiness` summary with `headlineClaimsAllowed=false`, `missingRequirementCount=7`, and missing requirements grouped by owner/artifact. The package-action artifact now carries `streaming-matrix-boundary` so one representative streaming writer proof cannot be described as full streaming parity without performance-owner approval or broader matrix evidence. This keeps local proof usable while making stronger headline copy a deliberate product decision rather than an accidental interpretation of a digest.

## Do Not Promote Yet

- Formula safe rename or edit-producing rename.
- Formula workbook-context `prepareRename` for defined names, table names, table columns, sheet refs, 3D refs, spill refs, or external refs.
- Formula language-service release copy beyond corpus-backed rejection-first primitives.
- Columnar sidecars as a product feature.
- Release proof bundle as signed provenance.
- Compact proof-report digests in the release index before artifact storage, privacy filtering, and stable canonicalization are owner-approved for both top claims.
- Universal Excel formula compatibility.
- Collaborative/CRDT claims from retained viewport patches.
- Claims backed only by private workbook corpora.
