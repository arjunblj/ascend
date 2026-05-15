# 2026-05 Release Claim Board

Date: 2026-05-15

## Purpose

Freeze the next release claims into proof-shaped language. This board is deliberately conservative: it states what Ascend may say today, what proof is still missing, and which owner loop should carry the next step. Research should not keep promoting new production surfaces while these proof gaps remain open.

Current stewardship rule: hand off only the top one or two product claims to implementation loops. Everything else stays as proof packaging, validation, or "do not promote yet" until the missing evidence is explicit.

Latest proof refresh, 2026-05-15T19:20:27Z:

- `release-proof-index --no-timings --owner-handoffs-json` reports `releaseGate=blocked-by-publication-policy`, `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, and `missingRequirementCount=9`.
- The ranked 10-direction portfolio remains stable: ranks 1 and 2 are the only top implementation handoffs; ranks 3 through 6 are proof-packaging-only; ranks 7 through 10 remain do-not-promote-yet.
- Safe-open proof covers 9 cases with 6 public fixtures, 2 generated edge packages, 1 malformed package, 8 OK, 1 rejected, 4 review-before-hydration routes, and stable shape `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`.
- Package-action proof covers 8 cases with 4 public fixtures, 2 generated workbooks, 2 generated edge packages, action totals `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1`, and stable shape `b4cf0755f2fd40ff577ab87b6600ee14a58bbd69608796c903759e7e8b4d29e8`.
- Formula-assist proof remains rejection-first only: 1685 public formulas sampled, 2322 reference spans, 25 binding roles, 3 LET-local prepare-rename OK targets, and 1692 prepare-rename refusals. No edit-producing rename is allowed.
- Release packageability evidence is now local-tarball-smoke-backed for SDK and CLI/API/MCP apps, but remains below publication, provenance, API listener lifecycle, MCP stdio protocol, and retention/privacy claims.

Latest proof refresh, 2026-05-15T10:15:42Z:

- `release-proof-index --no-timings --json` still reports `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, release gate `blocked-by-publication-policy`, and 9 missing requirements across product, correctness, performance, and release owners.
- Safe-open compact proof covers 9 cases with 6 public fixtures, 2 generated edge packages, 1 malformed package, 4 review-before-hydration routes, malformed rejection, and stable shape `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`. Signed and unknown-part evidence remain generated structural packages.
- Package-action compact proof covers all five action classes with 8 cases, 4 public fixtures, 2 generated workbooks, 2 generated edge packages, action totals `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1`, one representative streaming proof, and stable shape `b4cf0755f2fd40ff577ab87b6600ee14a58bbd69608796c903759e7e8b4d29e8`. Generated edge-package policy, unsupported-feature boundaries, provenance wording, and streaming-matrix wording remain owner gates.
- Formula-assist proof remains rejection-first only: 1685 public formulas sampled, 2322 reference spans, 25 binding roles, 3 LET-local prepare-rename OK targets, and 1692 prepare-rename refusals. The latest proof was run with `--no-timings`; no latency claim is promoted. No edit-producing rename is allowed.
- Practical latency phase profiles are useful performance diagnostics and stay excluded from release proof artifacts until a tracked-clean release-environment run over approved public inputs exists.
- A cluster of tiny casing fixes landed for journals, verifier checks, chart/pivot/connection selectors, hyperlink locations, and chart/pivot ownership. Treat these as correctness hygiene under the auditable mutation claim, not as new release surfaces.
- `release-proof-index` now tags each next owner action with a `nextStepKind`, so top blockers are visibly owner decisions, validation runs, optional harness expansion, or publication policy rather than implicit permission for research to add new surfaces.
- `release-proof-index` now also reports `implementationSurfacePromotionAllowed=false` while these blockers remain, with boundary text that unresolved proof requirements do not authorize new SDK, CLI, API, or MCP surfaces.
- `release-proof-index` now emits `implementationHandoffs` for exactly the top two product-shaped claims, carrying owner loops, proof commands, blocker IDs, next-step kinds, and the same fail-closed surface-promotion flag.
- Each `implementationHandoff` now carries `proofRequired`: fixture, benchmark, surface, validation gate, competitor contrast, honest boundary, and kill criterion. This keeps the claim ladder product-shaped in the machine-readable release artifact, not just in prose.
- `release-proof-index` now emits `deferredClaims` for lower-ranked directions, so formula rename, columnar sidecars, oracle routing, viewport history, token-bounded agent view, and agent traces stay out of top implementation handoff until their proof gaps close.
- `practical-latency-contracts --input-preset public-tracked` can now generate a tracked-harness edit workbook for edit/verify phases while labeling it as generated evidence. This improves performance-owner reproducibility, but practical latency remains excluded from release proof until owner-approved tracked-clean threshold wording exists.
- `agent-phase-profile` now records normal and prepared commit output workbook byte medians alongside JSON payload bytes, giving performance owners a clearer edit/verify profile without promoting latency or size thresholds.
- Current owner-handoff rerun keeps the gate fail-closed: `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, 9 missing owner requirements, and exactly two implementation handoffs: safe unknown workbook opening and auditable package-part mutation.
- `agent-phase-profile --progress --timeout-ms` can now stream phase-event NDJSON while keeping final JSON on stdout, giving performance owners partial progress evidence for longer local probes without promoting latency claims.
- Blocker grouping from `release-proof-index --json`: product has 2 fixture-decision blockers, correctness has 1 boundary-approval blocker, performance has 1 validation-run blocker plus 1 streaming-boundary decision, and release owns 4 publication-policy blockers.
- Journal exactness tests now assert generated lossy journal issues use exported v1 issue-code, surface, and reason registries before classification, strengthening auditable mutation proof without adding a new surface.
- Safe-open public fixture scan now uses the tracked git XLSX/XLSM corpus and still finds no signed/unknown replacement: 223 public fixtures scanned, 1 protected fixture rejected, 0 matches. Product must accept disclosed generated structural fixtures or explicitly acquire public binaries.
- Package-action compact proof has 4 public fixtures, 2 generated workbooks, and 2 generated edge packages with all five action classes covered; the generated edge packages should support only disclosed structural package-action wording, not provenance, trust, or full streaming-parity claims.
- Compact safe-open and package-action reports omit workbook bytes, input/output byte fields, per-part digest fields, and generated artifacts. Treat them as owner-review pointers until release approves storage, privacy filtering, and canonicalization policy.
- Worksheet reader metadata extraction now uses markup-compatibility-stripped XML for custom sheet views, raw extension lists, and controls; the SDK trust moat also uses a real imported shared-formula fixture for corruption blocking. Treat this as correctness hygiene under preservation-first/auditable mutation, not as claim promotion.
- Compact agent plan checks and blocking write-policy diagnostics now have regression coverage for exact imported shared-formula refs (`Label!A3`, `Label!A2`) in the real shared-formula trust fixture.
- Safe-open local latency rerun with `--repeat 3 --warmup 1` produced public-fixture open-plan medians from 0.133 ms to 2.514 ms and full/open-plan ratios from 12.35x to 88.98x. Keep `release-latency-run` missing until performance approves the release environment, input set, repeat policy, and non-threshold wording.
- Blocked agent commits now restore the caller's in-memory workbook when post-apply write-policy checks throw; the real shared-formula trust moat asserts no speculative `Label!C1` edit remains after the blocked commit. Treat this as correctness hygiene under auditable mutation, not a new claim.
- Package-action streaming boundary rerun still shows `streamingProofCases=1` and `streamingRegenerateParts=1`; allowed wording remains one representative streaming dirty-sheet proof, not full streaming parity across add/drop/error or macro/chart cases.
- API and MCP write errors now preserve structured failed-apply journal evidence in error details, including `JOURNAL_BUILD_FAILED`, surface `package-parts`, reason `journal-build-failed`, and undo policy `build-failed`.
- `release-proof-index --owner-handoffs-json` now includes `nextOwnerActions` with `acceptanceEvidence` and `forbiddenShortcut` for every missing top-claim gate. This makes owner approval actionable without changing SDK, CLI, API, or MCP surfaces.
- The human-readable release proof index now renders those owner actions as a Markdown table. Compact safe-open and package-action reports remain unchanged; the table is release-index reporting only.
- Each top `implementationHandoff` now carries its own cloned `blockingActions`, so safe-open and package-action owners can consume a self-contained handoff without cross-referencing the global owner-action list.
- Package-action `docprops-passthrough` now uses checked-in public `fixtures/xlsx/calamine/date_1904.xlsx` instead of a generated edge package. Package-action proof source counts are now 4 public fixtures, 2 generated workbooks, and 2 generated edge packages; the remaining generated edge packages are signature invalidation and unknown-part error, so the fixture-policy gate stays missing.
- `package-action-fixture-scan --json` now scans the tracked git XLSX/XLSM corpus for package-action replacement candidates. Current result: 223 fixtures scanned, 1 rejected, `docPropsCore=191`, `docPropsCustom=16`, `calcChain=52`, `customXml=4`, `macro=2`, `chartOrDrawing=46`, `signaturePackage=0`, and `syntheticUnknownPathFamily=0`.
- `release-proof-index --owner-handoffs-json` now includes `fixturePolicy`: generated structural fixture acceptance criteria, public-binary-required criteria, tracked fixture scan commands, current generated structural cases, and external policy/provenance references. The policy is an owner-decision aid only; it does not satisfy `public-edge-fixtures` or `edge-fixture-policy`.
- `fixturePolicyEvidence` now records the latest tracked fixture scan summaries in the owner handoff: safe-open scanned 223 fixtures with 0 signature/unknown replacements, package-action scanned 223 fixtures with `signaturePackage=0` and `syntheticUnknownPathFamily=0`, and `ownerApprovalRequired=true`.
- `fixturePolicy.approvalChecklist` now lists four pending product/release decisions with validation commands, acceptance evidence, and rejection conditions. All checklist items remain `pending-owner-decision`; no release gate is satisfied by the checklist itself.
- `performancePolicy.approvalChecklist` now lists pending performance decisions for `safe-open-proof/release-latency-run` and `package-action-proof/streaming-matrix-boundary`. It records validation commands and rejection conditions while keeping local timing below release-threshold wording and one streaming case below parity wording.
- `streamingMatrixEvidence` now records the concrete streaming boundary in owner-handoff JSON: one representative streaming proof case, covered action kinds `passthrough` and `regenerate`, missing action kinds `add`, `drop`, and `error`, and public non-streaming cases `calc-chain-drop`, `macro-passthrough`, and `chart-sidecar-accounting`. This is proof of the boundary, not satisfaction of `streaming-matrix-boundary`.
- `generatedFixtureDecisionEvidence` now records each generated structural fixture case with tracked replacement evidence and owner-only allowed/forbidden use. Product can review safe-open `signed`, `unknown-part`, `malformed`, and package-action `signature-invalidation-drop`, `unknown-part-error` without treating generated bytes as public binaries.
- `compactReportPublicationEvidence.policyDecisions` now records the four release-owned publication decisions for compact reports: artifact storage path, retention/privacy filtering, canonicalization subject, and offline verification expectations. Compact report digests remain unpublished and non-attestation wording remains mandatory.
- `correctnessBoundaryEvidence` now reports `missingFeatureNames` and `ownerEscalationRequired` so unsupported-feature proof regressions become explicit owner blockers, not just a failed broad boolean.
- `readiness.claimBlockerBoard` now groups release blockers by product-shaped claim and owner loop. It is derived from missing `readyWhen` gates, so it is a routing view, not a second source of truth.
- `safeOpenLatencyValidationEvidence` now records the release-latency gate as performance-owner evidence. The default owner handoff is untimed and explicitly forbids release or threshold claims until performance approves a tracked timed run, input set, repeat/warmup policy, and non-threshold wording.
- `claimPortfolio` now records the ranked 10-direction portfolio in owner-handoff JSON. It marks only safe unknown workbook opening and auditable package-part mutation as top implementation handoffs; formula primitives, token-bounded agent view, retained viewport patch history, and release proof bundle are proof-packaging-only; formula oracle routing, property journal laws, columnar scan sidecars, and agent workflow observability are do-not-promote-yet.
- Release packaging remains a proof-bundle blocker, not a research surface invitation. Current source manifests are private/workspace-based, app builds emit declarations only, CLI/API/MCP dist JS is placeholder output, CLI still points `bin.ascend` at source and depends on TUI, API/MCP lack publish manifests or bins, and SDK agent docs resolve from repo layout instead of package assets. Product/release should own the next packaging harness.
- Shared-formula journal coverage was tightened for structural formula rewrites: `moveRange` formula-surface restoration now uses edit preimages so every rewritten shared-formula member is reported in journal issues. Treat this as correctness hygiene under auditable package-part mutation, not a new claim.
- Practical latency target selection now carries p95/CV and stability fields for key first-view, edit-verify, and repeated-inspection phases. This helps performance owners reject noisy median targets before production work, but remains diagnostic routing and does not satisfy release latency gates.
- `release:sdk:smoke` now proves a narrow SDK-only temp external install from a built tarball without consumer overrides. It verifies create/open/plan/commit/reopen/check/recalc from `@ascend/sdk`; it does not close CLI/API/MCP packaging, bundled docs, publication policy, provenance, or signed artifact claims.
- CLI/API/MCP packageability now has a narrow implementation proof: app entrypoints can be imported without auto-running, `build:js` emits app JS artifacts, app dist manifests include `ascend`, `ascend-api`, and `ascend-mcp` bin entries, and SDK agent docs are copied into the SDK release artifact. This does not yet prove app tarball installation from an external consumer.
- Legacy array formula integrity now rejects occupied detached cells inside an array formula range as `legacy-array-range-member-mismatch`. Treat this as correctness hygiene under auditable mutation. It does not change the formula-intelligence board, and it does not authorize rename.
- `release:apps:smoke` now proves temp external installation of packed CLI/API/MCP tarballs. It runs the installed `ascend` bin, executes an installed API capabilities request, and calls installed MCP capabilities tool/resource callbacks without workspace dependencies. This closes the app-tarball-install and basic runtime-smoke proof gap but not publication policy, signed provenance, artifact retention, API listener lifecycle, or a real MCP stdio protocol session.
- Practical latency attribution now ranks measured shared plan sub-phases ahead of aggregate prepared-plan timing. The latest public-tracked edit/verify dry run selected prepared-output reopen as the largest phase, but remains diagnostic-only because it was one sample with generated edit input and a dirty harness.
- `release-proof-index` now exposes `releasePackageabilityEvidence` in full JSON, owner-handoff JSON, and Markdown. It routes SDK/app tarball smoke commands and missing publication/protocol requirements to release owners while keeping `headlineClaimsAllowed=false` and `implementationSurfacePromotionAllowed=false`.

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
| **Formula language-service primitives:** Ascend exposes formula IDE building blocks: parse diagnostics, token/reference spans, hover, completions, reference cycling, binding roles, and a corpus-backed rejection-first prepare-rename guard. Allowed wording must say the guard refuses workbook-context and reference targets; it must not imply edit-producing rename. | **Fixture:** SDK rejection matrix covers LET declarations/uses with shadowing, workbook-context names, table names, table columns, structured item selectors, external workbook refs, 3D refs, spill refs, literals, punctuation, and parse failures. The proof harness samples public POI/ClosedXML formula workbooks.<br>**Benchmark:** `fixtures/benchmarks/formula-assist-proof.ts` currently discovers 1685 public formulas, samples 1685, reports 2322 reference spans, 25 binding roles, 3 LET-local prepare-rename OK targets, and 1692 prepare-rename refusals by reason. The current stewardship proof was run with `--no-timings`, so no latency wording is promoted.<br>**Surface:** existing SDK `formulaAssist`, CLI formula assist, API formula-assist, and MCP formula-assist only; no new rename surface.<br>**Validation gate:** parser/span/binding-role tests, cross-surface assist tests, and the formula-assist proof harness. Latest local proof passed; no production code was changed in this stewardship refresh.<br>**Competitor contrast:** HyperFormula owns formula engine/dependency graph breadth; Ascend should claim workbook-preserving edit intelligence and refusal semantics, not a broader formula IDE.<br>**Honest boundary:** **no edit-producing rename**, no safe workbook-context rename, no table/defined-name rename from formula assist, and no claim that all formula references can be rewritten safely. | Product/DX stewardship only. Proof/spec boundary; not an implementation handoff. |
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

Latest no-rename surface audit:

```bash
rg -n "prepareRename|renameTarget|formulaAssist|formula[-_ ]assist|rename target|occurrenceRanges|apply.*rename|WorkspaceEdit|TextEdit" packages apps fixtures research/experiments/syntheses -g '*.ts' -g '*.md'
bun test packages/sdk/src/formula-edit.test.ts fixtures/benchmarks/formula-assist-proof.test.ts
bun test apps/cli/src/cli.test.ts -t "formula assist"
bun test apps/api/src/server.test.ts -t "formula-assist"
bun test apps/mcp/src/index.test.ts -t "formula_assist"
```

Result: the current surfaces expose `renameTarget` metadata and formula-local LET prepare evidence only. They do not expose a formula rename command, workbook-wide edit plan, LSP `WorkspaceEdit`, or operation-owned rename application path. Existing workbook operations such as sheet/table rename are separate workbook mutation operations, not formula-assist rename.

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
| Auditable package-part mutation | `package-action-proof` command, stable shape digest `b4cf0755f2fd40ff577ab87b6600ee14a58bbd69608796c903759e7e8b4d29e8`, all five action classes, one representative streaming writer proof, journal issue refs, post-write audit status, validation commands, and accepted or resolved `readyWhen` gates | Do not publish headline copy if synthetic edge packages are hidden, if chart XML is called byte-passthrough, if one streaming proof is described as full streaming matrix coverage, or if the proof implies SLSA/in-toto/signed provenance |

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

Machine-readable handoff: `releaseProofOwnerHandoffIndex.fixturePolicy` now carries the same policy for product/release owners. `releaseProofOwnerHandoffIndex.fixturePolicyEvidence` carries current tracked-corpus scan evidence for the top fixture gates. Current generated structural cases are `safe-open-proof=signed,unknown-part,malformed` and `package-action-proof=signature-invalidation-drop,unknown-part-error`. Current decision remains `owner-approval-required`. Its approval checklist has four pending items: `safe-open-proof/public-edge-fixtures`, `package-action-proof/edge-fixture-policy`, `safe-open-proof/publication-boundary`, and `package-action-proof/provenance-boundary`.

Current application: the safe-open signed and unknown-part cases can remain generated local proof because they exercise OPC package topology and are disclosed as synthetic. They should still block stronger release copy until accepted or replaced; `fixtures/benchmarks/safe-open-fixture-scan.ts` currently finds no tracked public binary replacement.

Product fixture gate recommendation: resolve the current top fixture gates by explicitly accepting disclosed generated fixtures for structural package-topology proof, not by waiting indefinitely for public signed/unknown binary replacements. Public binaries should remain required before wording that depends on real-world workbook authoring behavior, vendor-specific semantics, user trust, malware scanning, signed provenance, or Excel repair behavior. Latest local scan: `fixtures/benchmarks/safe-open-fixture-scan.ts --json` scanned 223 tracked public XLSX/XLSM fixtures, rejected 1, and found 0 signed/unknown replacement candidates. Package-action proof currently has 4 public fixture cases, 2 generated workbook cases, and 2 generated edge-package cases; the remaining generated edge cases are signature invalidation drop and unknown-part error.

Safe-open gate resolution attempt: as of the 2026-05-15 clean tracked-tree probe, the public-edge-fixture gate cannot be resolved by replacement from the tracked public corpus. `bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json` scans git-tracked fixtures, skips ignored external/stress fixture folders for filesystem fallback, scanned 223 fixtures, rejected 1 protected fixture, and found 0 signed/unknown-part replacements. A constrained external probe found one public unknown-part candidate, `node-projects/excelForge/src/test/Book 1.xlsx`, that routes to `metadata-only` review with risk family `preservedOther`, part count 50, relationship count 37, and SHA-256 `9c5426fa71ff68cc7e40e19e02b5992daf91da5754ef643d2db2f89bd70bb122`. Do not vendor or treat it as a release fixture yet: the GitHub repository API reports `license: null`, the source is not checked in here, and this does not solve the signed-workbook case. Web research found stable documentation for OPC signatures and unknown relationships, plus signing APIs/articles, but no durable public signed-XLSX binary fixture suitable to vendor into the release proof. The practical resolution path is therefore product approval of disclosed generated structural fixtures or an explicit fixture-acquisition policy, not silent promotion to public-binary evidence. A timed local run of `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json` produced public-fixture open-plan medians from 0.133 ms to 2.514 ms and full/open-plan ratios from 12.35x to 88.98x, plus review routing for macro, ActiveX, synthetic signature, and synthetic unknown-part inputs; this is useful diagnostic evidence, but the release-latency gate should remain missing until a performance owner approves environment, repeat count, inputs, and threshold wording.

Safe-open owner acceptance checklist: current proof has 9 cases: 6 public fixtures, 2 generated edge-package fixtures, and 1 generated malformed package. Generated/disclosed cases are `signed`, `unknown-part`, and `malformed`; 4 cases route to review before hydration; malformed rejection is true. Product must either accept disclosed generated structural fixtures, approve acquisition/vendor rules for the external unknown-part candidate, or replace the generated cases with tracked public binaries. Performance must run tracked-clean release-environment latency evidence and approve any threshold wording. Release must approve boundaries that exclude malware scanning, sandboxing, file trust, active-content safety, signed provenance, and malformed-package recovery. Until those checkboxes are complete, `safe-open-proof` remains `headlineClaimAllowed=false`.

Package-action unsupported-feature boundary matrix: correctness can approve narrower wording without changing mutation surfaces if each unsupported feature is described as package accounting, not semantic support.

| Feature boundary | Current proof evidence | Allowed wording | Forbidden wording |
| --- | --- | --- | --- |
| Digital signatures | `signature-invalidation-drop` generated package records signature origin/signature XML as dropped after edit. OPC says signatures validate package components, but signer identity and origin validation are consumer responsibilities. | "Ascend detects signature package parts and reports invalidation/drop evidence when a write changes the package." | "Ascend preserves, verifies, re-signs, or attests signatures." |
| Calc chain | `calc-chain-drop` now uses public `fixtures/xlsx/poi/Booleans.xlsx` and records `xl/calcChain.xml` as dropped for formula topology edits. Microsoft describes calc chain as dynamic calculation-order metadata maintained by Excel. | "Ascend reports calc-chain drop/regeneration decisions when edits make cached calculation order unsafe." | "Ascend proves Excel recalculation equivalence or cached formula freshness." |
| Chart and drawing sidecars | `chart-sidecar-accounting` public fixture reports drawing/media/style/color sidecars as passthrough while generated workbook/styles/sheet/sharedStrings parts change. | "Ascend accounts for chart/drawing sidecars separately from regenerated workbook parts." | "Chart XML is byte-passthrough" or "Ascend semantically understands every chart feature." |
| Macros and ActiveX | `macro-passthrough` public XLSM fixture reports macro-bearing parts as package evidence; safe-open routes macro/ActiveX to review. | "Ascend records macro/ActiveX package preservation and review routing." | "Macros or ActiveX are safe, sandboxed, scanned, or executable through Ascend." |
| Unknown package parts | `unknown-part-error` generated package records one `error` action and failed post-write audit. OOXML permits unknown relationships, but producers are not required to round-trip them. | "Ascend can fail closed with an explicit unknown-part error." | "Ascend preserves or understands arbitrary unknown parts." |
| Streaming writer | `docprops-passthrough` has one representative streaming proof with one regenerated worksheet part and passthrough-byte equality count. | "One representative streaming writer package-action proof exists." | "Full streaming parity across all package-action scenarios." |

Boundary decision: keep `unsupported-feature-boundary` missing in `release-proof-index` until a correctness owner explicitly accepts the allowed/forbidden wording above. The matrix is enough to make the next approval concrete, but it should not silently flip the gate to satisfied.

Machine-readable handoff: `releaseProofOwnerHandoffIndex.correctnessPolicy` now carries this unsupported-feature matrix plus one pending approval checklist item for `package-action-proof/unsupported-feature-boundary`. `releaseProofOwnerHandoffIndex.correctnessBoundaryEvidence` now proves every current matrix row has backing local evidence from package-action and safe-open proof cases. Current decision remains `owner-approval-required`; accepting the checklist would approve wording only, not semantic support for signatures, chart XML, cached formula freshness, macro/ActiveX safety, unknown parts, or full streaming parity.

Package-action owner acceptance checklist: current proof has 8 cases: 4 public fixtures, 2 generated workbook cases, and 2 generated edge-package cases. Disclosed generated cases are `regenerate-existing-sheet`, `add-sheet-part`, `signature-invalidation-drop`, and `unknown-part-error`; all five action classes are present with `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, and `error=1`; source graph evidence and journal package issues appear in every case; `unknown-part-error` is the only post-write audit failure; one representative streaming proof exists. Product must accept or replace remaining generated edge packages. Correctness must approve unsupported-feature boundaries and journal/package issue compatibility. Performance must accept narrow streaming wording or expand the matrix. Release must approve non-provenance wording and compact report publication policy. Until those checkboxes are complete, `package-action-proof` remains `headlineClaimAllowed=false`.

Package-action provenance boundary audit: release can approve local digest wording only if it stays below the attestation threshold used by SLSA, in-toto, GitHub artifact attestations, and Sigstore. Current local probe result: `release-proof-index` reports `signed=false`, `attestation=false`, `headlineClaimsAllowed=false`, and keeps `package-action-proof/provenance-boundary` as missing; the package compact report omits workbook bytes and per-part artifact digests and says it is not signed provenance, SLSA, in-toto, or third-party attestation. That is the right fail-closed shape. It is proof of reproducible local evidence, not proof of build origin, signer identity, transparency-log inclusion, or consumer policy verification.

| Boundary | Current proof evidence | Allowed wording | Forbidden wording |
| --- | --- | --- | --- |
| Local shape digest | `release-proof-index` records artifact SHA-256 and stable-shape SHA-256 for local proof JSON. | "The release proof index identifies local evidence artifacts by digest and stable shape." | "Digest evidence is tamper-evident provenance." |
| Signed provenance | `signed=false`, `attestation=false`, and `provenance-boundary` remains missing. | "No signed provenance is produced by this proof harness." | "SLSA provenance", "signed proof bundle", or "certified build origin." |
| in-toto statement | No `_type`, `subject`, `predicateType`, predicate, DSSE envelope, or signer policy is produced. | "Future attestation work could map proof artifacts to an in-toto/SLSA subject after release policy exists." | "This proof is an in-toto attestation." |
| GitHub/Sigstore attestation | No `actions/attest`, OIDC identity, Fulcio certificate, Rekor inclusion proof, or cosign bundle is generated. | "Attestation is out of scope for the current local proof index." | "GitHub/Sigstore verified" or "transparency-log backed." |
| Publication policy | `compact-report-publication-policy` remains missing for both top artifacts. | "Compact report commands are reproducibility pointers." | "Compact report digests are publishable release artifacts before storage, privacy filtering, and canonicalization are approved." |

Boundary decision: keep `provenance-boundary` missing until a release owner explicitly approves the local-proof wording above or implements a real attestation pipeline. Do not add package-action code for this.

Package-action edge fixture policy audit: the tracked-corpus ZIP-entry scan covers 223 git-tracked XLSX/XLSM fixtures and shows partial public replacement potential, not full replacement. The corpus has 52 workbooks with `xl/calcChain.xml`, 16 with `docProps/custom.xml`, 4 with `customXml/` parts, 2 with `vbaProject.bin`, and 46 with chart/drawing parts. It has 0 with `_xmlsignatures/` package signatures and 0 with the synthetic unknown path family used by `unknown-part-error`. The harness now uses public `fixtures/xlsx/poi/Booleans.xlsx` for `calc-chain-drop` and public `fixtures/xlsx/calamine/date_1904.xlsx` for `docprops-passthrough`. The docProps public case records `passthrough=4`, `regenerate=5`, `drop=0`, `add=0`, `error=0`, keeps one representative streaming proof, and preserves `docProps/core.xml`, `docProps/app.xml`, and `docProps/custom.xml` as byte-equal passthrough parts in the local probe. `signature-invalidation-drop` and `unknown-part-error` still need generated fixtures or a fixture acquisition policy.

Boundary decision: keep `edge-fixture-policy` missing. The calc-chain and docProps public replacements improve fixture quality, but they do not close the product gate because signed and unknown-part edge cases remain unreplaced.

Package-action streaming matrix boundary audit: performance can approve only narrow wording today. Current local proof covers 8 package-action cases overall and all five action classes, but streaming proof exists only for `docprops-passthrough`; that streaming case covers `passthrough` and `regenerate`, regenerates only `xl/worksheets/sheet1.xml`, and preserves 3 passthrough byte digests. The seven non-streaming package-action cases are `regenerate-existing-sheet`, `add-sheet-part`, `calc-chain-drop`, `signature-invalidation-drop`, `macro-passthrough`, `chart-sidecar-accounting`, and `unknown-part-error`.

| Streaming boundary | Current proof evidence | Allowed wording | Forbidden wording |
| --- | --- | --- | --- |
| Representative streaming passthrough | One public docProps fixture has streaming proof with `passthrough` and `regenerate` actions. | "A representative streaming dirty-sheet write preserves passthrough parts while regenerating the dirty worksheet." | "Streaming package-action proof covers every action class." |
| Add/drop/error actions | Add, drop, and error are proven only in non-streaming package-action cases. | "Non-streaming proof covers add/drop/error package actions." | "Streaming add/drop/error parity is proven." |
| Macro/chart public fixtures | Public macro and chart fixtures are non-streaming in the current package-action proof. | "Public macro/chart package accounting is proven in the standard writer path." | "Streaming macro/chart preservation is release-proven." |
| ZIP streaming semantics | JSZip and spreadsheet libraries document streaming/write-only tradeoffs and ordering constraints. Ascend has its own streaming writer tests, but this release proof matrix has one case. | "Streaming wording is limited to the tested dirty-sheet passthrough case." | "Streaming mode is semantically equivalent across all workbook/package features." |

Boundary decision: keep `streaming-matrix-boundary` missing until a performance owner either accepts one representative case as sufficient for release wording or expands the proof matrix. The safe release phrasing is "one representative streaming proof exists," not "full streaming parity."

Compact-report publication policy audit: both top artifacts now have compact report commands, but compact report digests are intentionally not indexed. Current local probe confirms `safe-open-proof` and `package-action-proof` each have a `compactReportCommand`, each carry `compact-report-publication-policy(missing,release)`, neither artifact includes a compact digest field, and neither compact report embeds workbook bytes. This is the correct pre-publication state because RFC 8785-style canonicalization, artifact storage, retention/privacy filtering, and offline verification expectations are release policy decisions, not research defaults.

Machine-readable handoff: `releaseProofOwnerHandoffIndex.compactReportPublicationEvidence` now records compact report command presence, JSON byte sizes, top-level fields, forbidden payload field scan results, and the missing publication policy requirements. Current result: compact report digests indexed `false`, forbidden payload fields embedded `false`, `generatedAt` included `true`, and owner approval required `true`.

| Publication boundary | Current proof evidence | Allowed wording | Forbidden wording |
| --- | --- | --- | --- |
| Compact command pointer | Release index includes compact report commands for both top artifacts. | "Compact report commands reproduce claim-safe summaries locally." | "Compact reports are published release evidence artifacts." |
| Compact digest indexing | No compact report digest is indexed for either top artifact. | "Digest publication is deferred until storage and canonicalization policy exists." | "Compact report digests are stable release commitments." |
| Workbook byte minimization | Compact safe-open and package-action reports omit workbook bytes and full proof artifacts in the local probe. | "Compact reports are minimized summaries." | "Compact reports are privacy-reviewed artifacts." |
| Canonicalization | Stable-shape digests exist for full proof artifacts; no shared compact-report canonicalization policy exists. | "Full proof artifacts have stable-shape digests." | "Compact report bytes are canonical or signer-ready." |

Boundary decision: keep `compact-report-publication-policy` missing for both top artifacts until a release owner defines storage path, retention/privacy filtering, canonicalization, and verification expectations. Compact commands are useful proof pointers; compact report digests are not release artifacts yet.

Release-index enforcement: `fixtures/benchmarks/release-proof-index.ts` marks both current top artifacts as `headlineClaimAllowed: false` with `releaseGate: blocked-by-publication-policy`. It exposes per-artifact `readyWhen` requirements by owner loop, compact report reproduction commands, and an aggregate `readiness` summary with `headlineClaimsAllowed=false`, `missingRequirementCount=9`, and missing requirements grouped by owner/artifact. The package-action artifact now carries `streaming-matrix-boundary` so one representative streaming writer proof cannot be described as full streaming parity without performance-owner approval or broader matrix evidence. Both top artifacts carry `compact-report-publication-policy`, so compact report commands stay discoverability pointers only and compact report digests stay out of the index until artifact storage, privacy filtering, and canonicalization policy are owner-approved. This keeps local proof usable while making stronger headline copy a deliberate product decision rather than an accidental interpretation of a digest.

Owner-action acceptance evidence: each missing gate now has a machine-readable acceptance target and forbidden shortcut. Product fixture gates must accept disclosed generated structural packages or replace them with approved public binaries. Performance gates must use tracked-clean release-environment inputs and cannot use private or dirty local timing. Correctness gates must approve allowed/forbidden unsupported-feature wording and cannot imply chart byte passthrough, signature verification, Excel-fresh cached formulas, or unknown-part understanding. Release gates must keep local proof below SLSA, in-toto, Sigstore, GitHub artifact attestation, signed-provenance, and tamper-evident storage thresholds unless a real attestation pipeline is implemented.

Markdown handoff: `releaseProofIndexMarkdown` now includes a `Next Owner Actions` table with rank, artifact, gate, owner loop, priority, next step, acceptance evidence, and forbidden shortcut. This is for human release review only and does not make compact reports publishable.

Per-claim handoff data: `releaseProofOwnerHandoffIndex` now includes both the global `nextOwnerActions` queue and per-handoff `blockingActions`. The per-handoff actions are cloned, not shared references, so downstream consumers can inspect or transform one handoff without mutating the global queue.

Owner-action priority: `readiness.nextOwnerActions` ranks fixture disclosure/replacement first, correctness boundaries second, performance evidence third, streaming wording fourth, publication wording fifth, and compact report publication policy last. Compact report policy still blocks digest publication, but public/generated fixture policy and claim-boundary approval decide whether the top claims are credible.

## Do Not Promote Yet

- Formula safe rename or edit-producing rename.
- Formula workbook-context `prepareRename` for defined names, table names, table columns, sheet refs, 3D refs, spill refs, or external refs.
- Formula language-service release copy beyond corpus-backed rejection-first primitives.
- Columnar sidecars as a product feature.
- Release proof bundle as signed provenance.
- Compact proof-report digests in the release index before artifact storage, privacy filtering, and stable canonicalization are owner-approved for both top claims.
- A shared compact-report canonicalization helper before release owners define the artifact subject, storage location, and verification expectations.
- Universal Excel formula compatibility.
- Collaborative/CRDT claims from retained viewport patches.
- Claims backed only by private workbook corpora.
