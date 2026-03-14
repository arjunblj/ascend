# Open-Source Spreadsheet Engine Landscape: Technical Comparison

A comprehensive technical analysis for building the fastest, most efficient spreadsheet engine in TypeScript. Research conducted March 2025.

---

## Executive Summary

| Engine | Language | Storage | Formula | Key Strength | Key Weakness |
|--------|----------|---------|---------|---------------|--------------|
| **SheetJS** | JS | Sparse (default) / Dense | None | Ubiquity, format support | Full-file load, V8 string limit |
| **ExcelJS** | JS | In-memory | Basic | Streaming write API | Memory leaks in streaming, no backpressure |
| **openpyxl** | Python | Lazy (read-only) | None | Read-only/write-only modes | 50× memory in standard mode |
| **Calamine** | Rust | Range-based | None | Fast parse, multi-format | 85% time in worksheet layer |
| **Handsontable** | JS | Virtualized DOM | HyperFormula | Row/col virtualization | Screen reader limits |
| **Luckysheet** | JS | Store-based | Built-in | Collaborative, Excel-like | Monolithic, 90% JS |
| **Univer** | TS | DI/microkernel | Lexer→Parser→Interpreter | Lambda, supertables | Complex DI |
| **IronCalc** | Rust | Model-based | 300+ functions | WASM, multi-locale | No collaboration yet |
| **HyperFormula** | TS | Sparse/Dense policy | ~400 functions | Dependency graph, GPU | JS GC overhead |
| **Formualizer** | Rust | Arrow-backed | 320+ functions | 4× faster than HF, 4× less memory | Benchmarks in progress |

**Bottom line**: Rust/WASM engines (Formualizer, IronCalc) show 4× calculation speed and 4× lower memory vs pure JS. For TypeScript-native, the biggest wins come from: (1) lazy evaluation context (GRID: 10% gain), (2) sparse/dense policy (HyperFormula), (3) formula compilation (Ascend codegen), (4) dependency-graph incremental recalc.

---

## 1. SheetJS (xlsx)

### Architecture
- **Storage**: Sparse by default (object keyed by cell address); dense mode (`dense: true`) uses array-of-arrays.
- **Read**: Full-file load into memory. No streaming read for XLSX (ZIP requires random access).
- **Write**: Streaming export to CSV/JSON/HTML/XLML only; XLSX write is in-memory.

### Performance
- Dense mode significantly faster for large sheets; V8 favors arrays over large objects.
- Streaming write: `to_csv`, `to_json`, `to_html`, `to_xlml` for memory-efficient export.
- Web Workers recommended for large files in browser.

### Limitations (GitHub issues)
- **V8 string length limit**: Hardcoded; cannot be bypassed with `--max-old-space-size`. Affects large XLSX.
- **Browser OOM**: Crashes at ~45MB (300K rows) to 250MB (500K+ rows) without dense mode.
- **XLSX streaming**: Impossible by format; entire file must be in memory.
- Pro Edit offers surgical XML modification for append on large files.

### Techniques to Steal
- Dense mode as default for large sheets; sparse for sparse data.
- Streaming export pattern: read once (dense), stream out CSV/JSON.
- Web Worker offload for parse + stream in browser.

### What They Do Poorly
- No formula engine; read/write only.
- Sparse default hurts large-file users who don't know about `dense: true`.
- No incremental or partial read.

---

## 2. ExcelJS

### Architecture
- **Streaming**: `stream.xlsx.WorkbookWriter` for streaming write.
- **Memory**: In-memory workbook model; streaming only affects write path.

### Performance
- Streaming write can still hit OOM with millions of rows (issues #709, #2916).
- No built-in backpressure (issue #734); unbounded memory if write can't keep up.
- Tips: call `.commit()` often, disable `useStyles`/`useSharedStrings` when possible.

### Formula Support
- Formulas in `f` field, A1-style (no leading `=`).
- No documented formula evaluation; streaming behavior for formulas unclear.

### Limitations
- Memory leak in `WorkbookWriter` (issue #2916).
- No Web Streams support yet (issue #2753).
- Backpressure handling missing.

### Techniques to Steal
- Streaming write API pattern.
- Explicit `.commit()` for row flushing.

### What They Do Poorly
- Streaming promises that don't hold under load.
- No formula engine; formula support is structural only.

---

## 3. OpenPyXL (Python)

### Architecture
- **Read-only mode**: `load_workbook(filename, read_only=True)` — lazy loading, near-constant memory.
- **Write-only mode**: `Workbook(write_only=True)` — append rows, ~10MB even for huge data (requires lxml).
- **Standard mode**: Full load; memory ≈ 50× file size (e.g. 2.5GB for 50MB file).

### Performance
- Optimized modes: ~0.15s faster per 1000×50 rows.
- Unoptimized: 15–33 rows/sec on 20K+ row files; degrades as it progresses.

### Optimization Pitfalls
- Avoid repeated `sheet.cell(row, col).value` — cache in variables.
- Don't rely on `max_row`/`max_column` (can include empties); use `iter_rows()`.
- Don't call `sheet.cell()` repeatedly — assign to temp vars.
- Parallel `ProcessPoolExecutor` on same file causes I/O bottlenecks.

### Techniques to Steal
- **Read-only**: Lazy cell iteration, explicit `close()`.
- **Write-only**: Append-only, single save, minimal memory.
- Clear separation of modes; document memory implications.

### What They Do Poorly
- Standard mode memory explosion (50×) is easy to hit.
- No formula evaluation.

---

## 4. Calamine (Rust)

### Architecture
- Pure Rust; supports xls, xlsx, xlsm, xlsb, ods.
- Uses `quick_xml` for XML; Serde for deserialization.
- `Range` for cell areas; `RangeDeserializer` for typed conversion.

### Performance
- 21MB / 270K rows: under a minute with `--release`, often <10s.
- **Bottleneck**: Worksheet reading ~85% of parse time (~22s for large files).
- `quick_xml` baseline for 1.3GB worksheet XML: ~6s → calamine adds ~16s overhead.
- Rewrite improved 22s → ~13s via: size-optimized strings, structure matching dimensions.

### Techniques to Steal
- Size-optimized strings for cell values.
- Match structure to actual dimensions (avoid overallocation).
- Focus optimization on worksheet layer, not just XML parse.

### What They Do Poorly
- Read-only; no formula engine.
- Significant overhead over raw XML parsing.

---

## 5. Handsontable

### Architecture
- **Virtualization**: Row and column virtualization by default; only visible viewport in DOM.
- **Renderers**: 10 built-in; renderer = full cell DOM. Prefer `valueFormatter` for simple transforms.
- **Viewport**: `viewportRowRenderingOffset`, `viewportColumnRenderingOffset` for pre-render outside view.

### Performance
- Refactor of viewport calculators: up to **3×** improvement.
- Tips: constant row/column sizes, disable `autoRowSize`/`autoColumnSize`, limit CSS animations.

### Limitations
- Virtualization breaks screen reader row counts.
- Browser search limited to visible cells.
- Commercial license for many features.

### Techniques to Steal
- Row/column virtualization with configurable pre-render offset.
- `valueFormatter` vs full renderer for performance.
- Constant dimensions to avoid layout thrash.

### What They Do Poorly
- Accessibility tradeoffs with virtualization.
- Tight coupling to commercial product.

---

## 6. Luckysheet / Univer (Chinese OSS)

### Luckysheet
- **Stack**: ~90% JavaScript; controllers, utilities, global modules.
- **Features**: Formatting, formulas, pivot, charts, comments, collaborative editing.
- **Collaboration**: WebSocket-based; history for undo/redo; MongoDB example backend.
- **Ecosystem**: Luckyexcel (import/export), chartMix, Vue/React/Node integrations.

### Univer
- **Architecture**: Microkernel + DI; 300K+ LOC, 297 modules.
- **Formula engine**: Lexer → Parser → Interpreter; dependency analysis; supports lambda, supertables.
- **Runs in**: Browser, Web Workers, Node.js.
- **Formula design**: LexerNode tree → AstNode; postfix for evaluation; `ArrayValueObject` with Numpy-like slice/filter for VLOOKUP/XLOOKUP.
- **Dependency**: Marks dirty cells, outputs execution queue; handles INDIRECT/OFFSET via pre-computation.
- **Decimal**: Uses decimal.js for precision.

### Techniques to Steal
- Lexer → Parser → Interpreter pipeline.
- `ArrayValueObject` for matrix ops; reverse index for iteration.
- `requestImmediateMacroTask` to avoid 4ms setTimeout limit.
- Feature registration (pivot, conditional format) via `getDirtyData`.

### What They Do Poorly
- Luckysheet: Monolithic, less modular.
- Univer: Heavy DI, steep learning curve.

---

## 7. IronCalc

### Architecture
- **Rust**: `ironcalc_base` (Model, formula eval), expressions, cell, worksheet, formatter.
- **WASM**: Runs in browser; minimal deps.
- **Formats**: roxmltree, zip for xlsx; bitcode/serde for serialization.

### Features
- 300+ Excel-compatible functions; multi-locale, timezone-aware.
- Import/export xlsx; multiple sheets, named ranges, dynamic arrays.
- Conditional formatting, custom formats.
- Bindings: Rust, JS, Python, Node.

### Techniques to Steal
- Rust + WASM for calculation worker.
- Modular crate layout (base, expressions, cell, worksheet).

### What They Do Poorly
- No real-time collaboration yet.
- No charts.
- Less public benchmark data than Formualizer.

---

## 8. HyperFormula

### Architecture
- **Headless**: Parser + evaluator; ~400 functions.
- **Dependency graph**: Only recomputes affected cells; cycle detection.
- **Address mapping**: `AlwaysSparse`, `AlwaysDense`, or `DenseSparseChooseBasedOnThreshold` (fill ratio).
- **Optimizations**: Succinct range deps, lazy CRUD updates, compressed repetitive formulas, GPU for MMULT/MAXPOOL/MEDIANPOOL, column index for VLOOKUP/MATCH.

### Performance
- v2.6.0: **60%** improvement from dependency graph work.
- `useColumnIndex` for VLOOKUP/MATCH on large unsorted data.
- `suspendEvaluation` / `resumeEvaluation` or batch for multi-op changes.

### Techniques to Steal
- Sparse/dense policy based on fill ratio.
- Column index for lookup functions.
- Batch operations to suspend recalc.
- Succinct range dependency representation.

### What They Do Poorly
- Pure JS: GC pauses, ~4× slower than Rust/WASM in benchmarks.
- ~80MB for 100K cells vs Formualizer's ~20MB.

---

## Cross-Cutting Topics

### Fastest Spreadsheet Engine Benchmarks
- **Formualizer vs HyperFormula** (10K rows, 100 formulas): Formualizer ~50ms, HyperFormula ~200ms.
- **100K cells**: Formualizer ~20MB, HyperFormula ~80MB.
- **Startup**: Formualizer ~1ms, HyperFormula ~10ms.
- **Excel/Sheets/Calc** (academic): Fail 500ms interactivity at 6K–150 rows (formula-heavy); sorting triggers full recompute; conditional format breaks ~80K rows.

### WASM Performance
- **Google Sheets**: JS engine was >3× slower than original Java. Migrated to WasmGC; after optimizations, **~2× faster than JS** (4× from initial WasmGC).
- **Key WasmGC optimizations**: Speculative inlining, devirtualization (~40% faster); use browser `RegExp` instead of re2j in WASM (~100× on regex); avoid JS-idiomatic array/map blur.
- **Formualizer**: Rust→WASM, ~2MB bundle; avoids GC; 4× faster, 4× less memory vs HyperFormula.

### Terminal Spreadsheet TUI
- **VisiData**: Structured data (Pandas-like); columns as lenses; `@asyncthread` for long iterators; lazy dynamic columns; progress objects; handles millions of rows with responsiveness over raw speed.
- **sc-im**: Curses-based, vim-like; ncurses UI.
- **Takeaway**: Async threads + progress, lazy computation, structured column model.

### Formula Compilation
- **Abacus Formula Compiler (AFC)**: Compiles formulas to bytecode; faster than interpretation.
- **GRID**: Fixed cost of `_makeCalcCellEvaluationContext` was 12.5% of recalc; reduced via lazy getters and shared `CellEvaluator` → **~10%** overall gain.
- **Ascend**: Codegen for IF/IFERROR/IFNA/SUM/VLOOKUP/MATCH/INDEX; cache of 4096 compiled fns.

### Incremental Computation
- **Dependency graph**: Standard; topological order for eval.
- **Range optimization**: Large ranges → quadratic deps; HyperFormula decomposes and reuses nodes.
- **Jane Street Incremental**: Self-adjusting computations; dynamic graph.
- **DataSpread**: Async formula computation; compress dep graph; return control quickly.

---

## Recommendations for Ascend

### Already Aligned
- **Sparse grid**: Chunked sparse storage (Ascend `SparseGrid`) matches best practice.
- **Formula codegen**: Compiled hot paths (IF, SUM, VLOOKUP, etc.) reduce interpreter overhead.
- **Dependency graph**: Incremental recalc of affected cells only.

### High-Impact Additions
1. **Lazy evaluation context** (from GRID): Single shared context with getters; avoid per-cell object creation.
2. **Sparse/dense policy** (from HyperFormula): Choose by fill ratio; dense for dense rectangles.
3. **WASM calculation worker** (from Google Sheets, Formualizer): Offload heavy calc to Rust/WASM; 2–4× speedup.
4. **Streaming export** (from SheetJS): CSV/JSON stream out without full materialization.
5. **Read-only / write-only modes** (from openpyxl): For large-file workflows, document memory implications.

### Avoid
- JS-idiomatic array/map blur if targeting WASM.
- Eager evaluation context construction.
- Ignoring V8 string limits for large XLSX (use dense, stream export).
- Full recalc on every change (dependency graph is essential).

---

## Sources

- SheetJS: docs.sheetjs.com, GitHub #2782, #2707, #1295, #2757, #138
- ExcelJS: GitHub #734, #2916, #709, #2753
- openpyxl: readthedocs optimized/performance, blog.dchidell.com, Stack Overflow
- Calamine: GitHub #145, #372, docs.rs
- Handsontable: handsontable.com/docs, GitHub #11069
- Luckysheet: dream-num.github.io, GitHub dream-num/Luckysheet
- Univer: docs.univer.ai (architecture, formula)
- IronCalc: ironcalc.com, GitHub, docs.rs
- HyperFormula: hyperformula.handsontable.com, GitHub #1027
- Formualizer: formualizer.dev, BSWEN comparison, Hacker News
- Google Sheets WasmGC: web.dev/case-studies/google-sheets-wasmgc
- GRID: alexharri.com/blog/grid-engine-performance
- SpreadsheetBench: spreadsheetbench.github.io
- VisiData: visidata.org, forked-visidata.readthedocs.io
- Formula compilation: formulacompiler.org, Berkeley DataSpread paper
- Incremental: HyperFormula dep graph, Jane Street Incremental, Adapton
