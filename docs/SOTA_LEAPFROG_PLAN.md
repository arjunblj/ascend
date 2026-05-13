# Ascend SOTA Leapfrog Plan

## Objective

Make Ascend the strongest open-source spreadsheet automation platform for:

- read/write performance
- formula correctness and recalculation speed
- XLSX/XLSM semantic fidelity and preservation
- real-workbook robustness
- developer experience
- agent/MCP workflows

The bar is not "good TypeScript spreadsheet library." The bar is beating the top OSS implementations in their strongest categories, with reproducible head-to-head evidence.

## Competitive Field

Ascend must be evaluated by category, because no single OSS project dominates every spreadsheet workload.

| Project | Category To Beat | Current Strength |
| --- | --- | --- |
| SheetJS CE | JS/TS workbook extraction and broad format conversion | Simple data conversion APIs, broad format familiarity, fast common reads |
| ExcelJS | JS workbook authoring and streaming-style API | Ergonomic workbook/worksheet/cell API, styling, streaming read/write |
| HyperFormula | Headless TS formula engine | Mature formula engine, CRUD, undo/redo, sorting, 400+ functions |
| Univer | AI-native spreadsheet app/framework | Full-stack spreadsheet system, agent positioning, 500+ functions, plugin ecosystem |
| IronCalc | Rust/WASM spreadsheet engine | Modern Rust engine, xlsx import/export, multi-language bindings, strong roadmap |
| Apache POI | JVM Excel compatibility baseline | Broad Office format support, XLS/XLSX, SXSSF memory-optimized large writes |
| openpyxl | Python XLSX/XLSM read/write | Dominant Python workbook manipulation library, familiar API |
| XlsxWriter | XLSX write fidelity | Rich write-only XLSX feature coverage, charts/images/macros, low dependency surface |
| Calamine | Rust read performance and breadth | Pure Rust read path for Excel/OpenDocument, borrowed-data APIs |
| rust_xlsxwriter | Rust write performance | Fast XLSX writer with constant-memory and native zlib options |
| libxlsxwriter | C write performance baseline | Native XLSX writer, constant-memory mode, source of XlsxWriter/rust_xlsxwriter comparison shapes |
| OpenXLSX / pyopenxlsx | C++/Python bulk read/write | Recent bulk-operation benchmarks and possible C++ speed ceiling for generated workbook IO |
| ClosedXML | .NET developer ergonomics | Intuitive OpenXML wrapper, tables, pivots, formula calculation docs |
| Excelize | Go XLSX/XLSM/XLSB automation | Pure Go read/write with streaming APIs for large worksheets |
| PhpSpreadsheet | PHP workbook automation | Maintained PHP successor to PHPExcel with broad spreadsheet format support |
| Formula.js | Function catalog comparison | Public function coverage list and JS scalar formula implementations |

## Where Ascend Can Win

Ascend should not copy any single competitor. The product advantage is the combination:

- Headless workbook automation like SheetJS/openpyxl.
- A real formula/dependency engine like HyperFormula/IronCalc.
- Fidelity and preservation like serious XLSX tools.
- Rich semantic verification and tracing that most competitors do not provide.
- First-class agent workflows that start with inspect/profile/read, not raw cell dumps.
- Stable machine contracts across SDK, CLI, API, and MCP.

The winning claim should be:

> Ascend is the fastest and safest open-source way for agents and developers to inspect, modify, recalculate, verify, and preserve real Excel workbooks.

## Required Scoreboards

### 1. XLSX Read Scoreboard

Competitors:

- Ascend
- SheetJS
- ExcelJS
- openpyxl
- Calamine
- Apache POI
- ClosedXML

Workloads:

- dense values
- sparse wide windows
- shared strings heavy
- styles heavy
- formulas loaded but not calculated
- table-heavy sheets
- pivot/chart/macro files in preservation mode
- selected-sheet reads
- metadata-only reads
- repeated inspect/read/trace warm workflow

Metrics:

- cold open latency
- first range read latency
- warm repeated read latency
- cells/sec
- peak RSS
- retained RSS after GC
- parsed metadata completeness
- failure/skip reason

### 2. XLSX Write/Fidelity Scoreboard

Competitors:

- Ascend
- SheetJS
- ExcelJS
- XlsxWriter
- openpyxl
- Apache POI/SXSSF
- ClosedXML

Workloads:

- dense generated workbook
- wide generated workbook
- style-heavy generated workbook
- table/validation/comment/hyperlink workbook
- dirty single-sheet edit in a large workbook
- formula edit with calc-chain invalidation
- macro workbook no-op save
- chart/pivot/slicer preservation no-op save

Metrics:

- write latency
- cells/sec
- output file size
- peak RSS
- part preservation count
- exact package fingerprint when no-op save is expected
- semantic diff against original
- Excel/openpyxl/POI reopen success

### 3. Formula Engine Scoreboard

Competitors:

- Ascend
- HyperFormula
- Univer formula engine where headless benchmarkable
- IronCalc
- ClosedXML formula engine
- Formula.js for scalar function behavior

Workloads:

- financial model functions
- lookup-heavy sheets
- SUMIFS/COUNTIFS criteria-heavy sheets
- growing SUM ranges
- volatile formulas
- shared formulas
- dynamic arrays and spill behavior
- dirty localized edits
- whole-column/whole-row references
- defined names and structured references

Metrics:

- initial recalc latency
- localized edit recalc latency
- formulas/sec
- dependency graph build time
- dirty set size
- AST/cache hit rate
- Excel-ground-truth pass rate
- error compatibility and explainability

### 4. Agent Workflow Scoreboard

Competitors:

- Ascend MCP/CLI/API
- Univer MCP or agent-facing stack
- general XLSX libraries plus a simple agent adapter
- SpreadsheetBench baselines where comparable

Workloads:

- inspect workbook and identify relevant sheets/tables
- locate headers and values
- read compact data slices
- explain formula lineage
- preview a mutation
- apply mutation
- recalculate
- verify output
- export deliverable

Metrics:

- tool calls to success
- tokens returned per useful fact
- pass rate on SpreadsheetBench-style tasks
- structured error recovery rate
- preview-before-write compliance
- deterministic reproducibility

## Corpus Plan

Ascend needs a corpus that mixes vendorable fixtures, locally downloaded public files, and user-provided/private stress files.

Vendored or script-downloadable sources:

- Apache POI spreadsheet test-data
- XlsxWriter comparison fixtures
- SheetJS test files where license/source permits
- Calamine tests
- openpyxl fixtures where license/source permits
- SpreadsheetBench sample or full dataset if license and access allow
- EUSES spreadsheet corpus if license and redistribution allow
- public government/statistical workbooks
- SEC, census, and other public reporting XLSX files

Each corpus entry needs:

- source URL
- license
- vendorable flag
- benchmark tier
- assertion class
- risk class
- feature tags
- known unsupported features
- expected preservation behavior
- expected semantic behavior

## Implementation Roadmap

### Phase 0: Public Contract Parity

Goal: remove surface drift before benchmarking.

- Keep OpenAPI in sync with HTTP implementation.
- Keep MCP operation validation in sync with canonical operation schema.
- Add schema parity tests for SDK, MCP, and API.
- Make all machine outputs versioned or explicitly marked unversioned.

### Phase 1: Benchmark Harness That Can Prove Wins

Goal: one command can compare Ascend against installed competitors.

- Replace ad hoc competitive scripts with unified benchmark envelopes.
- Add competitor adapters with explicit capability tags.
- Record runtime versions and command/environment metadata.
- Emit JSON suitable for regression gating.
- Separate unsupported from failed from slower.
- Keep first-party JS/TS adapters in-process for the minimum bar: Ascend must beat SheetJS and ExcelJS on real read and preservation workloads.
- Use preloaded workbook bytes for every in-process JS/TS adapter so timings compare library parse/write work instead of mixing filesystem latency into only one competitor.
- Add external runner protocol for non-JS leaders so Python, Rust, JVM, .NET, Go, and PHP competitors feed the same `BenchmarkSuiteResult` schema instead of fragmenting evidence.
- Keep `bench:competitive-real` fast by default with a quick corpus, and use `--full-corpus` for the broader local real-workbook sweep across public, stress, macro, style, formula, and generated comparison fixtures.
- Add `bench:competitive-scoreboard` as the objective ranking layer. It consumes a benchmark JSON file, gates speed rankings on `rankingEligible`, ranks by fidelity tier before latency, and still reports the fastest eligible competitor when a lower-fidelity rewrite is faster than exact package preservation.
- Make leader assertions statistically aware when repeat samples exist: median winners are still shown, but noisy non-significant reversals do not fail the gate.
- Support focused sweeps with `bench:competitive-real --category read|roundtrip` so higher-repeat evidence can be gathered for noisy read or write subsets without rerunning unrelated expensive cases.

External runner protocol:

- Manifest flag: `--runner-manifest runners.json`.
- Manifest shape: `{ "name": "openpyxl", "command": ["python3", "fixtures/benchmarks/runners/openpyxl_runner.py"], "categories": ["read", "roundtrip"], "capabilities": { "xlsmRoundtrip": true } }`.
- Invocation: `<command...> --operation <read|roundtrip> --file <path> --json`.
- Generated write manifest flag: `bench:competitive-io --write-runner-manifest writers.json`; generated write runners receive `<command...> --operation write --rows N --cols N --workload <name> --repeat N --warmup N --json`.
- Output: JSON object or `{ "assertions": object }`, with primitive assertion values only.
- Internal timing extension: runners with `"capabilities": { "internalTiming": true }` receive `--repeat N --warmup N` and may return `{ "assertions": object, "samples": [{ "durationMs": number }] }`; those samples are used directly so process startup is not counted as library work.
- Capability tags: `valueOnlyRead` marks runners that intentionally validate cached/scalar values instead of formula text; `metadataOnlyRead` marks runners that support workbook metadata reads without hydrating sheet cells and can participate in generated `read-metadata-only` profiles.
- Timing note: one-shot external commands include process startup. Serious runners should provide a long-lived wrapper or enough repetitions inside the command to amortize startup, and results must be labeled `executionScope=external-process`.
- Memory note: parent Bun RSS/heap is not a competitor memory measurement. External runners must report their own peak/retained memory before those metrics can be compared.

Runner targets to build:

- JS/TS: SheetJS, ExcelJS, node-xlsx where useful.
- Python: openpyxl, XlsxWriter, PyExcelerate, fastexcel/Polars, WolfXL if it remains OSS and mature enough.
  - Current local runners: openpyxl for formula-preserving read, metadata-only read, and semantic roundtrip; OpenPyXL normal/write-only, XlsxWriter normal/constant-memory, and PyExcelerate for generated write workloads; fastexcel for Calamine-backed value-only read, expected to be disqualified on formula-fidelity workloads.
- Rust/C: Calamine, rust_xlsxwriter, libxlsxwriter.
- JVM: Apache POI XSSF and SXSSF.
- .NET: ClosedXML and raw OpenXML SDK where relevant.
- Go: Excelize.
- PHP: PhpSpreadsheet.

### Phase 2: Real Workbook Corpus Expansion

Goal: benchmark against real Excel workloads, not only synthetic sheets.

- Add downloader scripts for public fixture sources.
- Normalize corpus manifest tags.
- Add independent expected workbook shape metadata instead of treating Ascend-derived sheet/cell/formula counts as ground truth.
- Make correctness assertions gating inputs to rankings: invalid parse, unsupported operation, semantic mismatch, and fidelity loss must be status fields, not passive annotations.
- Emit `rankingEligible` only for correctness-passing cases (`pass`, `exact-package-match`, or `semantic-roundtrip-pass`) so speed comparisons cannot accidentally reward a wrong parse or unsupported rewrite.
- Track semantic value/formula cells separately from physical OOXML `<c>` nodes so read-performance rankings are not distorted by style-only blanks, while fidelity scoreboards still see physical preservation obligations.
- Add competitor capability tags at runtime so unsupported categories like macro-preserving `.xlsm` roundtrip are skipped explicitly instead of counted as failures or misleading rewrites.
- Add package-fidelity assertions.
- Separate exact-package identity, semantic workbook equivalence, relationship/part preservation, and Excel reopen validity; byte identity is a strong no-op-preservation win, but not the only fair fidelity measure.
- Add semantic assertions for tables, names, formulas, validations, comments, hyperlinks, pivots, slicers, and charts.
- Add Excel/openpyxl/POI reopen validation lanes where available.
- Keep fidelity and performance visible as separate facts: exact package identity beats semantic rewrite for no-op preservation, but the scoreboard must still surface any speed gap against the fastest semantic competitor as an optimization target.

### Phase 3: Read Path Dominance

Goal: beat read competitors on cold and warm workloads.

- Make `WorkbookSession` the first-class read unit for SDK, API, and MCP.
- Add explicit session IDs for API/MCP or document why path cache is enough.
- Add table-aware reads everywhere.
- Add column-pruned reads.
- Wire `changedSince` into MCP/API compact reads.
- Keep sparse windows proportional to populated cells, not rectangle area.

### Phase 4: Formula Engine Dominance

Goal: beat HyperFormula/IronCalc on targeted automation workloads while improving Excel compatibility.

- Persist compiled formula state per session.
- Expand Excel-ground-truth fixtures for top financial/modeling functions.
- Add range-aware dependency nodes and localized invalidation.
- Add richer formula diagnostics with precedent values and evaluation traces.
- Benchmark dirty localized edits as a first-class category.

### Phase 5: Write/Fidelity Dominance

Goal: make preservation and semantic edits auditable and safer than competitors.

- Make preview-first mutation the default agent path.
- Add transactional edit planner for structural edits.
- Formalize part ownership and dirty-part patching.
- Expand no-op save and dirty edit fidelity contracts.
- Add inverse operations or undo journals for agent rollback.

### Phase 6: Agent SOTA

Goal: make Ascend the best spreadsheet MCP/API surface.

- Add `read_table`.
- Add stateless `eval`.
- Add MCP resources for sheet/table/name metadata.
- Add server instructions for inspect -> profile -> extract -> preview -> write -> verify.
- Add response size limits with automatic compact/TSV fallback.
- Add task-level evaluator for SpreadsheetBench-style workflows.

## Recent Evidence Log

### 2026-05-13: Values-Mode Numeric Hydration

- Bottleneck: dense raw-OOXML values reads still spent measurable time in per-cell grid writes after the byte worksheet parser avoided DOM parsing.
- Change: `SparseGrid` now has a contiguous plain-number span path, and the values byte parser batches only adjacent non-date numeric cells. One-cell numeric runs still use the old setter so mixed text sheets do not pay span overhead.
- Evidence: `dense-values`, raw OOXML, 65,536 x 10, `--phase read` improved from a prior 48.77 ms median to 45.94 ms median over 7 repeats after warmup.
- Guardrail: `fastexcel-reader-65536`, mixed 50% text, `--phase read` stayed in the recent 73-78 ms range at 73.95 ms median.
- Validation: focused core/reader tests, `bunx tsc --build`, Biome, and `bun run ci:perf-smoke` passed.
- Negative result: byte-backed and sequential shared-string resolvers reduced RSS only marginally and regressed the mixed-text read profile to 83 ms and 175 ms medians. Keep the existing lazy SST resolver until access-locality evidence supports a chunked index or other design.

### 2026-05-13: Capped First-Window Read Evidence

- Hypothesis: agent/TUI inspect workflows should not pay full-workbook hydration cost before showing the first compact window.
- Measurement: `xlsx-read-phase --phase capped-agent-window` now measures `readXlsx({ mode: "values", maxRows: 500 })` followed by the same compact window read as the full `agent-window` phase.
- Evidence: on `fastexcel-reader-65536`, capped first-window total was 13.75 ms median standalone and 15.01 ms median in `--phase all`, compared with 77.92 ms full open plus window in the same all-phase run.
- Product implication: SDK/API/MCP/TUI should expose a first-window or lazy-session open path that returns partial load metadata and pagination, then hydrates more rows/sheets on demand.

## Near-Term PR Sequence

1. Fix public contract drift and MCP operation parity.
2. Add schema parity tests.
3. Convert competitive benchmarks to the standard result envelope.
4. Add competitor adapters for SheetJS, ExcelJS, HyperFormula, and Calamine.
5. Add external runner implementations beyond the initial `openpyxl` runner.
6. Add corpus download/audit support for Apache POI and XlsxWriter fixtures.
7. Add `read_table` to SDK, CLI, API, and MCP.
8. Add session-backed API/MCP workflow benchmarks.

## Claim Criteria

Do not claim "SOTA" unless the repo can produce:

- reproducible head-to-head benchmark JSON
- competitor versions
- real workbook corpus manifest
- cold and warm results
- memory results
- correctness/fidelity assertions
- clear unsupported-surface accounting
- CI regression gates for Ascend-only benchmarks
- scheduled or manual competitive benchmark runs
