# Spreadsheet Engine Research: Architecture & Performance Deep Dive

Comparative technical analysis of major open-source spreadsheet/Excel engines, recalculation algorithms, and optimization techniques relevant to Ascend's engine design. Covers data structures, memory management, formula evaluation strategies, file I/O approaches, and cutting-edge techniques from both OSS and production engines.

---

## 1. SheetJS (xlsx)

**Language:** JavaScript | **Focus:** XLSX I/O (read/write), no formula evaluation | **Stars:** ~35k

### Cell Storage

SheetJS uses its **Common Spreadsheet Format (CSF)** — plain JS objects with no classes or prototypes. Two storage modes:

| Mode | Structure | When Faster |
|------|-----------|-------------|
| **Sparse** (default) | Object properties keyed by A1-style addresses: `sheet["A1"]` | Small worksheets, scattered writes (+10-15% faster assignment) |
| **Dense** (v0.19+) | Array of arrays: `sheet["!data"][R][C]` | Large worksheets in modern engines (+5-10% faster contiguous reads, +25% faster sorting via locality) |

Cell objects are plain `{t, v, w, f, ...}` — `t` = type tag, `v` = raw value, `w` = formatted text, `f` = formula string. No cell instances or class overhead.

### Performance Characteristics

- **No formula evaluation** — SheetJS is purely I/O. It preserves formula strings but never evaluates them.
- **Streaming export** via `XLSX.stream.to_csv`, `XLSX.stream.to_json`, etc. These are Node.js-compatible readable streams. No streaming *read* for XLSX (the entire ZIP must be decompressed).
- **Zip handling** is the dominant bottleneck. The library decompresses the entire XLSX ZIP into memory, then parses each XML file.
- **SharedStrings** parsing is sequential — every string in the workbook goes through a single shared string table lookup.

### Weaknesses

1. No streaming XLSX read. The entire file must be in memory (ArrayBuffer) before parsing begins.
2. No formula evaluation whatsoever.
3. Sparse mode's A1-key addressing creates V8 megamorphic property lookups on large sheets, defeating hidden class optimizations.
4. Style handling is minimal — styles are referenced by index but not deeply modeled.
5. Single-threaded, no Web Worker parallelism built-in.

### Worth Stealing

- **Dense mode's simplicity.** Array-of-arrays is the fastest JS structure for row-major iteration. Ascend's SparseGrid already uses a chunk-based approach that's more sophisticated, but the lesson holds: raw arrays beat Map/object for hot-path access.
- **Plain-object cell representation.** No class instances = lower GC pressure, compatible with structured clone / transfer.

---

## 2. ExcelJS

**Language:** JavaScript | **Focus:** XLSX read/write with rich API | **Stars:** ~14k

### Cell Storage

Hierarchical class-based model: `Workbook → Worksheet → Row → Cell`. Every cell is a class instance with properties for `_row`, `_column`, `_address`, `_value`, `style`, `_mergeCount`. This creates significant per-cell overhead.

### Memory Characteristics

- **Style duplication** is the critical problem. Each cell maintains its own style object instance even when styles are identical. WeakMap-based style caching actually makes things worse — users report OOM crashes with tens of thousands of rows, and style processing taking 20+ seconds for large files.
- Style object interning before write reduces `writeBuffer()` time by ~50%.
- Disabling the WeakMap entirely reduces memory pressure.

### Streaming Architecture

- **WorkbookReader** (streaming read): async iterator over worksheets, then async iterator over rows within each worksheet. Uses `node-unzipper` for ZIP extraction.
- **WorkbookWriter** (streaming write): sequential row appending, direct XML emission.
- Known bug: drops worksheets when processing files with >100 sheets due to inadequate stream draining in the unzipper.
- No Web Streams API support.

### Weaknesses

1. ~630 bytes/cell equivalent overhead from class instances (comparable to POI XSSF).
2. Style handling is an architectural failure — creates O(n) unique objects for n cells with the same style.
3. Streaming reader has reliability issues with many sheets.
4. Generator-based access is ~16x faster than individual cell access, but the API encourages the slow pattern.

### Worth Stealing

- **Negative lesson: style deduplication is critical.** Ascend's StyleRegistry approach (intern styles, reference by ID) is the right pattern. ExcelJS proves that per-cell style objects are catastrophic.
- **Async iterator streaming model** is clean API design for progressive read.

---

## 3. OpenPyXL (Python)

**Language:** Python | **Focus:** XLSX read/write | **Stars:** ~5k

### Memory Optimization Modes

| Mode | Memory | Speed | Constraint |
|------|--------|-------|------------|
| **Standard** | ~50x file size (e.g., 2.5 GB for 50 MB file) | Baseline | Full random access |
| **Read-Only** | Near-constant (~90% less) | ~2x faster | Forward-only iteration, `ReadOnlyCell` objects |
| **Write-Only** | <10 MB constant | ~2x faster | Append-only, single save |

### Streaming Architecture

- Read-only mode uses generators via `iter_rows()` with optional bounding parameters (`min_row`, `max_row`, `min_col`, `max_col`).
- Write-only mode writes XML directly without building an in-memory DOM.
- `lxml` (C-backed XML library) significantly accelerates write-only mode.
- Individual cell access via `sheet[row][col].value` is **16x slower** than generator-based iteration — a critical API design lesson.

### Weaknesses

1. Standard mode's 50x memory multiplier makes it unusable for large files without opt-in to restricted modes.
2. Read-only mode requires correct dimension metadata in the file.
3. Python's GIL prevents true parallel processing.
4. No formula evaluation.

### Worth Stealing

- **Bounded iteration API.** `iter_rows(min_row=5, max_row=100, min_col=2, max_col=10)` — efficient sub-range access without materializing the full grid. Relevant for Ascend's range-based formula evaluation.
- **Write-only direct XML emission** — skip the DOM, stream XML directly. Ascend's io-xlsx writer should evaluate this approach.

---

## 4. Apache POI (Java)

**Language:** Java | **Focus:** Full Excel format support | **The "gold standard"**

### Three Implementation Tiers

| Tier | Class | Memory per Cell | Use Case |
|------|-------|-----------------|----------|
| **HSSF** | HSSFWorkbook | ~77 bytes | .xls (BIFF8), <65k rows |
| **XSSF** | XSSFWorkbook | ~630 bytes | .xlsx, full DOM in memory |
| **SXSSF** | SXSSFWorkbook | Window-limited | .xlsx streaming write |

XSSF's 630 bytes/cell comes from keeping the full XML DOM (`CTCell` objects from XMLBeans) in memory. A 90 MB XLSX can require multiple GB of RAM.

### SXSSF Streaming Architecture

SXSSF maintains a **sliding window** of N rows in memory (configurable, default 100). Rows outside the window are flushed to temporary disk files as raw XML. This enables writing millions of rows with constant memory.

**Critical limitation:** SXSSF cannot evaluate formulas referencing flushed rows — throws `RowFlushedException`. Formulas that reference only within the current window work, but cross-window references are impossible.

Temporary XML files can grow enormous (1+ GB for 20 MB CSV input) unless gzip compression is enabled on the temp files.

### Formula Evaluation Engine

POI's `WorkbookEvaluator` is one of the most complete open-source formula engines:

- **~140 built-in functions** implemented.
- **Evaluation cache:** `WorkbookEvaluator` maintains a cache of intermediate cell values. Must be explicitly cleared via `clearAllCachedResultValues()` when cells are modified between evaluations.
- **`OperationEvaluationContext`** carries all state needed to evaluate a single operation: evaluator reference, workbook, sheet index, source cell coordinates, evaluation tracker.
- **`EvaluationTracker`** detects circular references during evaluation.
- **Tree-walking interpreter:** Formulas are parsed to a token tree (Ptg array, mimicking Excel's internal format), then walked recursively.

### Weaknesses

1. XML DOM storage (XSSF) is architecturally wasteful — 8x overhead compared to HSSF.
2. Formula evaluator is tree-walking, not compiled — significant overhead for hot formulas.
3. SXSSF's window model breaks formula evaluation across windows.
4. Java GC pressure from millions of short-lived cell objects.

### Worth Stealing

- **Evaluation cache with explicit invalidation.** Ascend's engine should cache intermediate formula results and invalidate precisely via the dependency graph.
- **OperationEvaluationContext pattern.** Minimal evaluation context passed to each operation — but GRID showed this should be a mutable singleton, not per-cell allocation (see GRID section below).
- **Ptg (parse-thing) token representation** — compact formula IR that mirrors Excel's internal format.

---

## 5. Calamine (Rust)

**Language:** Rust | **Focus:** XLSX/XLS/ODS *read-only* | **Downloads:** 5.9M+

### Cell Data Representation

```rust
enum Data {
    Int(i64),
    Float(f64),
    String(String),
    Bool(bool),
    DateTime(ExcelDateTime),
    DateTimeIso(String),
    DurationIso(String),
    Error(CellErrorType),
    Empty,
}
```

`Range<Data>` is the primary data container — a rectangular area backed by a `Vec<Data>`. This is the key design choice: data is stored in a flat, contiguous vector with row-major layout.

### Performance Analysis

Built on top of `quick_xml` for XML parsing. Current overhead analysis for a 1.3 GB file:

| Layer | Time | Notes |
|-------|------|-------|
| Raw `quick_xml` parse | ~6 sec | Baseline |
| Calamine worksheet read | ~22 sec → ~13 sec (WIP rewrite) | 3-4x → 2x overhead |

Overhead sources:
- Shared string table lookups (string deref for every cell referencing shared strings)
- `String` heap allocations for cell values
- Type coercion during parsing (XML text → typed Data variant)

### Optimization Opportunities Being Explored

1. **Small String Optimization (SSO)** via `smartstring` crate: 5.4% improvement (831ms → 786ms on 150k rows) by keeping short strings inline, reducing heap allocations.
2. **Decoupled cell value representation** from parsing — separate parsing step from typed storage.
3. **Multi-threaded worksheet reading** — XLSX worksheets are independent compressed entries, theoretically parallelizable. Blocked by Rust's borrow checker requiring architectural changes for multiple file handles.

### No formula evaluation. Read-only.

### Worth Stealing

- **Flat `Vec<Data>` with row-major layout.** Contiguous memory gives excellent cache behavior for row-iteration workloads. Ascend's chunk-based approach is more flexible but should ensure chunk-internal data is contiguous.
- **Enum-based cell value** — `Data` enum is compact (size = max variant + discriminant). Ascend's `CellValue` uses tagged objects; for the hot path, a Rust-style discriminated union in TypedArrays could be faster.
- **SSO for cell strings.** Most cell strings are short (names, labels, codes). Inline storage avoids heap allocation. Ascend's `StringTable` interning achieves a similar goal but through a different mechanism.

---

## 6. IronCalc (Rust)

**Language:** Rust | **Focus:** Full spreadsheet engine with formula evaluation, WASM | **License:** MIT/Apache 2.0

### Architecture

The `Model` struct is the central data structure. Cell storage uses `sheet_data` — a map from row numbers to column data (essentially `BTreeMap<u32, BTreeMap<u32, Cell>>`).

Formula handling:
- Formulas are parsed to AST nodes and stored in a `parsed_formulas` cache indexed by sheet.
- Cells reference formulas by index into this cache.
- ~192 Excel functions implemented.

### Formula Evaluation

- **Tree-walking interpreter** — evaluates parsed formula ASTs directly.
- No compilation/codegen/JIT.
- Dependency tracking through formula displacement — when cells move, the system iterates all cells and updates formula references using `to_string_displaced()`. This is O(n) in total cells, not O(affected cells).

### WASM Support

Core engine compiles to WASM via `wasm-bindgen`, running in browsers without dependencies. This is the primary deployment target alongside native Rust.

### Weaknesses

1. **BTreeMap-based cell storage** — logarithmic lookup per cell access. Fine for small sheets, but pointer-chasing through tree nodes defeats cache for large sheets.
2. **No incremental recalculation** — the roadmap mentions updating the evaluation algorithm with a "support graph," suggesting full recalc is the current model.
3. **O(n) formula displacement** — shifting a single row requires iterating every cell in the workbook.
4. **Limited function coverage** — 192 functions vs Excel's 500+.

### Worth Stealing

- **WASM-first architecture.** Proves that a Rust spreadsheet engine can run performantly in browsers via WASM. If Ascend ever considers a Rust core, WASM compilation is viable.
- **Parsed formula cache.** Separate formula storage from cell storage, reference by index. Ascend's approach of storing formula strings per-cell could benefit from a similar interning/caching layer for parsed ASTs.

---

## 7. Univer (TypeScript) — Successor to Luckysheet

**Language:** TypeScript | **Focus:** Full-featured modern spreadsheet engine | **Stars:** ~8k

Luckysheet is no longer maintained; its creators now develop Univer, a ground-up rewrite with a fundamentally different architecture.

### Architecture: Plugin-Based with DI

Everything in Univer is a plugin: `@univerjs/sheets` for core logic, `@univerjs/sheets-ui` for Canvas-based rendering, `@univerjs/sheets-formula` for the formula engine. Features like pivot tables, conditional formatting, and data validation are separate plugins that register with the formula engine via `IFeatureCalculationManagerService`.

The system uses a **Command/Mutation/Operation** pattern with dependency injection. Business logic flows through commands, mutations represent atomic state changes, and operations handle the I/O layer. This is architecturally similar to Ascend's operation/patch model but with more formal DI infrastructure.

### Formula Engine: Five-Layer Pipeline

Univer's formula engine is the most thoroughly documented modern OSS formula architecture:

| Layer | Responsibility |
|-------|---------------|
| **Lexer** | Tokenizes formula strings into `LexerNode` trees, converts infix to postfix notation |
| **Parser** | Transforms `LexerNode` trees to typed `AstNode` trees (ValueNode, RangeReference, LambdaNode, etc.) |
| **Interpreter** | Recursive AST execution with async support; handles INDIRECT/OFFSET as `ReferenceObject` returns |
| **Engine** | Dependency analysis, formula execution ordering, syntax/semantic analysis |
| **Service/Model** | Formula data storage, custom names, supertables, scheduling |

Key design: the interpreter uses a **NumPy-inspired matrix computation model**. All values flow through `BaseValueObject` and its subtypes:
- `ArrayValueObject` — supports slice, filter, broadcasting, and reverse-indexed iteration for VLOOKUP/XLOOKUP
- `PrimitiveValueObject` — `NumberValueObject`, `StringValueObject`, `BooleanValueObject`
- `LambdaValueObject` — enables higher-order functions (MAKEARRAY, REDUCE, MAP, SCAN)

Numeric precision uses `decimal.js` under the hood.

### Rendering: Canvas-Based

Unlike DOM-based spreadsheets, Univer renders to Canvas with a custom layout engine. This enables smooth scrolling of large sheets but requires reimplementing text rendering, selection, and input handling.

### Web Worker Formula Execution

Formula evaluation runs in a **Web Worker** to avoid blocking the UI thread. The worker communicates results back to the main thread asynchronously, with runtime status reporting (total formulas, completed count).

### Dependency Analysis

When a cell changes, the dependency module marks all transitive dependents as dirty and outputs an execution queue in topological order. For reference functions like INDIRECT/OFFSET, the module performs pre-computation during dependency analysis by calling the Lexer and Parser to resolve the reference range before scheduling.

### Feature Registration Pattern

Plugins like pivot tables register with the formula engine through `IFeatureCalculationManagerService`:
1. Register a dependency range and a `getDirtyData` callback
2. When the dependency range is marked dirty, `getDirtyData` executes the feature's internal computation
3. The callback returns dirty areas and temporary data for formulas that depend on the feature's output

This lets complex features participate in the recalculation graph without the formula engine knowing their internals.

### Weaknesses

1. Heavy plugin infrastructure adds complexity for simple use cases.
2. Canvas rendering means no native browser text selection, accessibility, or find-in-page.
3. Formula engine uses tree-walking interpretation — no compilation/codegen.
4. `decimal.js` for all numeric computation adds overhead for precision most users don't need.

### Worth Stealing

- **NumPy-style ArrayValueObject.** Broadcasting, slice, and filter operations on array values simplify function implementations dramatically. Univer's VLOOKUP is ~10 lines because array operations handle the heavy lifting.
- **Feature registration for dirty propagation.** Letting external features (pivot tables, conditional formatting) register dependency ranges and dirty callbacks is an elegant extension point. Ascend could use this pattern for user-defined computed views.
- **Web Worker formula isolation.** Running the formula engine in a Worker is a clean separation that prevents calc-heavy workbooks from blocking UI. Ascend should evaluate this for browser deployments.
- **Lambda/higher-order function support.** LAMBDA, LET, MAP, REDUCE, SCAN — these are the future of Excel formulas. Univer's `LambdaValueObject` pattern shows how to thread closures through the evaluation pipeline.

---

## 8. LibreOffice Calc (C++)  

**Language:** C++ | **Focus:** Full desktop spreadsheet | **The open-source reference**

### Cell Storage: `mdds::multi_type_vector`

The breakthrough data structure, designed by Kohei Yoshida. Instead of wrapping each cell value in a heap-allocated cell object:

```
Traditional:  Cell* → {type, value_ptr} → actual_value
              (scattered across heap, terrible cache behavior)

multi_type_vector:
  [NumberBlock: 5.0, 3.14, 2.71, ...]  ← contiguous doubles
  [EmptyBlock: 200 cells]               ← zero storage
  [StringBlock: "hello", "world", ...]   ← contiguous strings
  [NumberBlock: 1.0, 2.0, ...]           ← contiguous doubles
```

Values of the same type in consecutive rows are stored in contiguous arrays within typed blocks. This means:
- SUM over 100k numbers iterates a flat `double[]` — no indirection, perfect cache utilization.
- Empty regions cost zero memory (just a block header with a count).
- Block type transitions only create new blocks at type boundaries.

**Benchmark:** Iterating 100k × 1k cells for summation: **1.96s → 0.36s** (5.4x speedup) compared to the older heap-based approach.

Block lookup was originally linear (assuming sequential access), later changed to **binary search** to support threaded formula evaluation where access patterns are non-sequential.

### Dependency Tracking: Broadcaster-Listener Pattern

- **Single cell references:** `ScFormulaCell` registers as a listener on the referenced cell's broadcaster via `ScDocument::StartListeningCell()`. When the referenced cell changes, it broadcasts to all listeners.
- **Range references:** `ScBroadcastArea` objects represent range broadcasters. Multiple formula cells referencing the same range share a single broadcast area, avoiding O(n²) listener registrations.
- When a formula cell is deleted/modified, `EndListeningTo()` breaks the relationship.

### Threaded Calculation

Introduced by Collabora (Tor Lillqvist, Dennis Francis):

- **Formula groups** — sequences of formulas in a column with identical structure but offset references — are the unit of parallelism.
- Independent formula groups of the same length/weight execute on separate threads.
- Required making code thread-safe: static locals → thread-local storage, iterator indices moved out of class definitions.
- Speedup proportional to core count, limited by formula group independence.

### Weaknesses

1. C++ complexity — the codebase is enormous (~10M LOC across LibreOffice).
2. Broadcaster-listener pattern creates a web of pointers that's hard to reason about and difficult to parallelize beyond formula groups.
3. Block binary search adds overhead for random access patterns (compared to O(1) array index).

### Worth Stealing

- **`multi_type_vector` concept is the single most important optimization to study.** Contiguous typed blocks with run-length encoding of empty regions. Ascend's SparseGrid chunk system partially captures this (chunks are dense internally) but doesn't achieve full contiguity of same-typed values.
- **Formula group threading.** Identify structurally identical formulas in columns and evaluate them in parallel. Ascend's codegen could detect these patterns.
- **Shared broadcast areas for ranges.** Avoid O(n²) dependency edges for overlapping ranges — similar to HyperFormula's range node decomposition.

---

## 9. HyperFormula (TypeScript)

**Language:** TypeScript | **Focus:** Headless formula engine | **Stars:** ~2k

Included because it's the most relevant comparable to Ascend's formula engine — same language, similar goals.

### Dependency Graph Architecture

Three-phase pipeline: Parse → Build dependency graph → Evaluate in topological order.

**Range node optimization** — the key innovation:

When multiple cells depend on overlapping ranges (e.g., `B1=SUM(A1:A1)`, `B2=SUM(A1:A2)`, ..., `B100=SUM(A1:A100)`), naive implementation creates O(n²) dependency edges.

HyperFormula decomposes ranges: when encountering `B5:D20`, it checks if `B5:D19` already exists. If so, `B5:D20` = `B5:D19` + row 20 cells. For associative functions (SUM, MAX, COUNT), the result is computed incrementally from the sub-range result plus the new row.

### Formula Parsing

Uses **Chevrotain** parser (faster than Jison, PEG.js). Generates ASTs.

### Optimization Techniques

1. **Relative addressing for AST reuse.** Formulas `=A1+B1` in C1 and `=A2+B2` in C2 share the same AST with relative offsets. One parse, many cells.
2. **Lazy CRUD operations.** When cells are inserted/deleted, formula reference transformations are postponed until recalculation is actually needed.
3. **Topological sort** for evaluation ordering, with cycle detection.

### Performance

~200ms for 10,000 rows with 100 formulas (vs. Rust-based Formualizer at ~50ms). ~80 MB for 100k cells.

### Worth Stealing

- **Range decomposition for dependency graphs.** Ascend should implement this — it prevents O(n²) edge explosion for common patterns like running totals.
- **AST reuse via relative addressing.** Parse once, apply to entire column. Ascend's shared formula support is related but could be generalized.
- **Lazy structural transformations.** Don't update formula references on insert/delete — defer until recalc. Reduces the cost of structural edits from O(all formulas) to O(dirty formulas).

---

## 10. GRID (TypeScript)

**Language:** TypeScript | **Focus:** Commercial spreadsheet engine in browser

Not open-source, but their engineering blog reveals critical optimization insights.

### Key Finding: Fixed vs. Variable Recalculation Cost

In a spreadsheet with ~12,000 cells recalculating, **12.5% of time** was spent in `_makeCalcCellEvaluationContext` — a method that constructs the evaluation context object for each cell.

**Root cause:** Spreading `...this` (a 30+ property Workbook object) into a new object per cell evaluation. Plus eagerly computing properties that may never be used.

**Solution (yielded ~10% overall speedup):**

1. **Lazy getters** — compute properties only when actually accessed during evaluation.
2. **Single shared evaluation context** — create one `CellEvaluator` with a frozen evaluation context object. Update cell/ref via mutable private properties. Avoids per-cell object creation and GC pressure.

```typescript
class CellEvaluator {
  private evaluationContext: EvaluationContext; // created once
  private cell: Cell;
  private ref: Reference;

  evaluate(cell, ref) {
    this.cell = cell;   // mutate, don't recreate
    this.ref = ref;
    return evaluateAST(cell, this.evaluationContext);
  }
}
```

**Performance results:** Median -9.9%, weighted geometric mean -9.6% across their real-world test suite. Best case: 148ms → 43.5ms (-70.5%).

### Worth Stealing

- **Eliminate per-cell object allocation in the eval loop.** Ascend's `EvalContext` should be a mutable singleton, not created per cell. This is low-hanging fruit.
- **Lazy context properties via getters.** Only resolve sheet/workbook/mode when the formula actually needs it.
- **Benchmark against real workbooks**, not just synthetic tests. GRID runs ~200k unit tests + real document benchmarks on every build.

---

## 11. Recalculation Algorithms: The State of the Art

### Excel's Smart Recalculation (The Reference Implementation)

Excel's three-stage process:

1. **Dependency tree construction.** When formulas are entered/changed, Excel builds a tree of precedent→dependent relationships. From Excel 2007+, no limits on trackable dependencies (previously capped at 65,536).

2. **Calculation chain construction.** An ordered list of cells to evaluate. Excel scans top-to-bottom, left-to-right, building the chain. When it encounters a formula depending on an uncalculated cell, it reorders the chain (moves the dependent down).

3. **Dirty-flag propagation.** When a cell changes, Excel marks all direct and indirect dependents as "dirty." Only dirty cells are recalculated. Volatile functions (NOW, RAND, INDIRECT, OFFSET) are always marked dirty.

### Excel's Multithreaded Recalculation (MTR)

Introduced in Excel 2007:
- Up to **1024 concurrent threads** (default = number of cores).
- Excel identifies **independent branches** in the calculation chain that can execute in parallel.
- Thread-safe built-in functions run on worker threads; VBA/COM/unsafe functions run on main thread.
- Contention managed through careful memory partitioning.

### DataSpread Async Computation (Berkeley/UIUC Research)

Key insight: both dependency identification and optimal computation scheduling are **NP-Hard** problems for general spreadsheets.

Solution: **Lossy compressed dependency tables** that enable bounded-time dependency identification. Formulas compute in background with visual indicators for pending cells. The user can continue editing while computation proceeds.

This is the frontier of academic research on spreadsheet computation.

### Incremental Computation Frameworks

Two Rust frameworks directly applicable to spreadsheet recalculation:

**Adapton** — General-purpose incremental computing using a demanded computation graph (DCG). Tracks dependencies automatically, propagates changes minimally. Academic origin.

**Salsa** — Used in rust-analyzer. Defines programs as queries (pure functions from keys to values) with automatic memoization and incremental recomputation. More mature, wider adoption (2.8k stars).

Both implement the same core idea: when an input changes, only recompute the outputs that transitively depend on it, reusing cached results for everything else. This is exactly what a spreadsheet recalculation engine does.

---

## 12. Formula Compilation & JIT Techniques

### Approaches, Ordered by Performance

| Approach | Overhead per Eval | Startup Cost | Example |
|----------|-------------------|--------------|---------|
| **Tree-walking interpreter** | High (pointer chasing, virtual dispatch) | Lowest | Apache POI, IronCalc |
| **Stack machine / bytecode** | Medium (loop + switch dispatch) | Low | Excel's internal Ptg evaluation |
| **Closure-based codegen** | Low (direct function calls) | Medium | Ascend's current codegen.ts |
| **JIT to native** | Lowest (machine code) | Highest | Abacus Formula Compiler (Java bytecode) |

### Abacus Formula Compiler (Java)

The only open-source spreadsheet formula → bytecode compiler found:
- Compiles spreadsheet formulas to **JVM bytecode**.
- Generated classes implement a `Computation` interface.
- Optimization: **constant folding** with value caching during compilation.
- Eliminates runtime formula interpretation entirely.

### Ascend's Current Approach

Ascend's `codegen.ts` generates closure-based compiled functions for formulas involving IF, IFERROR, IFNA, and arithmetic/cell references. This is a good middle ground:
- Lower overhead than tree-walking (direct JS function calls, no AST traversal).
- No startup cost of JIT compilation.
- V8 can inline and optimize the generated closures.

### Opportunities

1. **Expand codegen coverage.** Currently limited to IF/IFERROR/IFNA + arithmetic + cell refs. Common functions like SUM, VLOOKUP, INDEX/MATCH could be codegen'd.
2. **Formula group compilation.** When N cells in a column share the same formula structure with offset references, generate a single function that takes a row offset parameter. Call it N times instead of generating N separate closures.
3. **SIMD for range aggregations.** SUM over a contiguous numeric range can be vectorized. With WASM SIMD, process 4 f64s per instruction.

---

## 13. WASM/SIMD Opportunities

### WASM SIMD for Spreadsheet Operations

Browser support: Chrome 91+, Firefox 89+, Safari 16.4+, Node 16.4+.

Applicable patterns for spreadsheets:

| Pattern | Application | Expected Speedup |
|---------|-------------|------------------|
| **Vectorized aggregation** | SUM, AVERAGE, MIN, MAX over numeric ranges | 2-4x |
| **Parallel comparison** | COUNTIF, SUMIF with numeric predicates | 2-4x |
| **Batch type coercion** | Number→string or string→number for columns | 2-3x |
| **Matrix operations** | MMULT, TRANSPOSE on numeric arrays | 4x+ |

WASM SIMD processes 128-bit lanes: 2×f64 or 4×f32 per instruction. For f64 spreadsheet values, this means 2x theoretical throughput for pure numeric operations.

### Key Constraint

SIMD only helps when data is **contiguous in memory and homogeneously typed**. This reinforces the value of:
- LibreOffice's `multi_type_vector` (contiguous typed blocks)
- Ascend's dense chunk representation (contiguous within chunks)
- Calamine's flat `Vec<Data>` layout

Scattered cell values in hash maps or tree structures cannot be SIMD'd.

### Practical WASM Strategy for Ascend

1. **Keep the TypeScript engine for logic/control flow.** Formula parsing, dependency graph, structural operations.
2. **WASM module for hot numeric paths.** SUM/AVERAGE/MIN/MAX over contiguous numeric ranges, compiled from Rust with `-msimd128`.
3. **SharedArrayBuffer for zero-copy data sharing** between TS engine and WASM module.
4. **Fallback to scalar TS** when data is non-contiguous or mixed-type.

---

## 14. Throbol: Compiled Spreadsheets with SIMD

**Language:** C++ | **Focus:** Real-time spreadsheet computation for robotics/control systems

Not a traditional spreadsheet library — Throbol is a compiled spreadsheet engine optimized for high-frequency recalculation (100-1000 Hz). Its techniques are relevant because they represent the performance ceiling for formula evaluation.

### Compilation Pipeline

1. **Parse** each cell into an AST independently (no cross-cell context needed for parsing)
2. **Type inference** — scan ASTs recursively, deduce types without annotations (except self-referencing cells like integrators)
3. **Emit code** in dependency order — traverses cells topologically, generating a sequence of callable operations

The emitted code is a `vector<function<void(CpuState&, Pad&)>>` — essentially a flat list of closures. Execution is a simple loop calling each function in order. No interpreter dispatch, no AST walking at runtime.

`std::function` overhead is ~10 machine instructions per opcode, so opcodes must be coarse-grained (a single opcode does meaningful work: matrix multiply, 3D primitive assembly, etc.).

### SIMD Batching

The key insight: since each spreadsheet recalculation uses **straight-line code with no branches**, multiple spreadsheet instances can execute in SIMD lockstep. Throbol runs 8 spreadsheet evaluations in parallel using SIMD:

- Values for each SIMD lane are allocated sequentially in a "pad" (scratchpad memory)
- The C++ compiler can see that loop iterations access sequential memory and auto-vectorizes to SIMD load/add/store instructions
- An `add` operation on 8 float values compiles down to ~8 ARM instructions: 4 loads, 2 SIMD adds, 2 stores

This yields 8x throughput for parameter sweeps (evaluating the same sheet with different inputs).

### Multicore

Worker threads take batches of up to 8 compatible jobs from a shared queue, combining SIMD and multicore for ~64x throughput on an 8-core machine.

### Performance

11.5 billion opcodes/second across all cores on an M1 MacBook Pro. 1M cell evaluations (1000 cells × 1000 timesteps) completes in under 15ms.

### Worth Stealing

- **Flat opcode sequence instead of tree-walking.** Ascend's codegen already moves in this direction, but Throbol shows the endgame: a flat `(state, scratch) => void` closure list is the fastest JS-native execution model. No dispatch, no AST nodes, just sequential function calls.
- **SIMD batching of independent evaluations.** When computing scenarios or sensitivity analysis (same formulas, different inputs), batch 2-4 evaluations together in WASM SIMD. Ascend's scenario/what-if features could benefit enormously.
- **Coarse-grained opcodes.** Make each compiled operation do enough work to amortize dispatch overhead. A SUM over 1000 cells should be one opcode, not 1000 individual adds.
- **Scratchpad memory model.** Pre-allocate a flat buffer for intermediate values, indexed by register number. Avoids per-cell heap allocation during evaluation.

---

## 15. Production Engines: How Google Sheets & Excel Differ from OSS

### Google Sheets: The WasmGC Migration

The most significant recent development in production spreadsheet engineering. Google Sheets ported its calculation engine from JavaScript to **WasmGC** (WebAssembly with garbage collection), achieving **2x faster calculations** than the JS version.

**History:**
- 2006: Launched with Java server-side calculation engine
- 2013: Moved calculation to the browser via GWT (Java → JS transpilation)
- Later: Switched to J2CL (Java to Closure JS transpiler)
- 2020: Chrome and Workspace teams began evaluating WasmGC
- 2021: First working WasmGC prototype
- 2022+: Optimization and rollout

**Why JS was too slow:** The JS calculation engine was >3x slower than the original Java server version. JS's loose types and dynamic behavior prevent JIT compilers from generating optimal code. TypeScript's types help developers but don't provide the guarantees compilers need for optimal codegen.

**Initial WasmGC performance was 2x *slower* than JS.** Key optimizations that brought it to 2x *faster*:

| Optimization | Impact | Notes |
|-------------|--------|-------|
| **Speculative inlining + devirtualization** | ~40% speedup | Replicating JVM optimizations that didn't exist for WasmGC yet |
| **Browser API delegation for regex** | ~100x for regex ops | `re2j` compiled to WasmGC → native `RegExp` API |
| **Platform-agnostic data structures** | Significant | JS-specific patterns (sparse arrays as maps) were slow on WasmGC |

**Key lesson for Ascend:** JavaScript-specific idioms (sparse arrays, dynamic property access, prototype chains) that V8 optimizes well are *anti-patterns* on other runtimes. If Ascend ever targets WASM, platform-agnostic data structures (typed arrays, flat buffers) will perform better across all targets.

**Architecture:** The calculation engine runs in a **Web Worker**, communicating with the main thread via `MessageChannel`. This is the same pattern Univer uses — further validation that formula evaluation should be isolated from the UI thread.

### Excel's Smart Recalculation: The Reference Implementation

Excel's recalculation engine is the most mature and optimized in existence. Three-stage process:

**Stage 1: Dependency Tree Construction**
When formulas are entered or modified, Excel builds a tree of precedent→dependent relationships. From Excel 2007+, no hard limit on trackable dependencies (previously capped at 65,536; exceeding this forced full recalculation).

**Stage 2: Calculation Chain Construction**
An ordered list of formulas to evaluate. Excel scans top-to-bottom, left-to-right. When it encounters a formula depending on an uncalculated cell, it doesn't just reorder — it **iteratively revises the chain** until no further improvement is possible. This dynamic reordering continues across multiple passes.

**Stage 3: Dirty-Flag Propagation**
On change, Excel marks all direct and indirect dependents as "dirty." Only dirty cells recalculate. Critically, **propagation continues even when a recalculated cell's value doesn't change** — the dependency walk is always exhaustive, only the actual computation is skipped for unchanged values.

**Multithreaded Recalculation (MTR):**
- Up to 1024 concurrent threads (default = core count)
- Independent branches of the calculation chain execute in parallel
- Thread-safe built-in functions run on workers; VBA/COM/unsafe functions run on main thread
- Memory partitioning prevents contention — each thread operates on disjoint cell regions

**Volatile Function Handling:**
Functions like NOW(), RAND(), INDIRECT(), OFFSET() are always marked dirty, forcing recalculation on every pass. This is a significant performance concern for workbooks with many volatile functions, as they create "dirty floods" that cascade through the dependency tree.

**Named Range Recalculation:**
Names are recalculated each time a formula referencing them is recalculated — they are not cached across formula evaluations within a single recalc pass. Unused names are never calculated.

### How Production Engines Differ from OSS

| Aspect | Production (Google/Excel) | Typical OSS |
|--------|--------------------------|-------------|
| **Recalc granularity** | Per-cell dirty tracking with chain reordering | Topological sort of dirty subgraph |
| **Threading** | 1024 threads, automatic branch parallelism | Single-threaded (most); worker threads (Univer) |
| **Formula coverage** | 500+ functions, 100% Excel compatibility | 140-400 functions, partial compatibility |
| **Memory management** | Sophisticated pooling, temp file spillover | In-memory only, OOM on large workbooks |
| **Validation** | Large corpus testing (thousands of real workbooks) | Unit tests on synthetic data |
| **Runtime** | WasmGC (Google), native C++ (Excel) | JS/TS tree-walking interpreters |
| **Undo/collaboration** | OT-based real-time merge | Command pattern or no support |

---

## 16. Common Performance Bottlenecks in Spreadsheet Engines

### The Fundamental Bottlenecks

**1. Dependency Graph Analysis (NP-Hard in General)**

Identifying which cells need recalculation when data changes is the core algorithmic challenge. Compressing the formula dependency graph to bound dependency identification time is NP-Hard for general spreadsheets (DataSpread, SIGMOD '19). In practice, engines use heuristics:
- Excel: iterative chain reordering with dirty-flag propagation
- HyperFormula: range decomposition to avoid O(n²) edges
- Formualizer: CSR edge format for cache-friendly traversal

**2. Per-Cell Overhead in the Evaluation Loop**

The single most common bottleneck across all engines. Creating objects per cell during evaluation dominates recalculation time for formula-heavy workbooks:
- GRID found 12.5% of recalc time spent constructing per-cell evaluation contexts
- ExcelJS spends ~630 bytes per cell on class instances
- Apache POI XSSF keeps full XML DOM per cell

The fix is universal: **shared mutable evaluation context**, not per-cell allocation.

**3. Style/Format Proliferation**

Style objects explode memory usage when not deduplicated:
- ExcelJS: OOM crashes with tens of thousands of rows due to per-cell style objects
- OpenPyXL standard mode: 50x file size memory multiplier, largely from style objects
- Solution: style interning (Ascend's StyleRegistry approach)

**4. Shared String Table Sequential Lookup**

XLSX files store strings in a shared string table. Every cell referencing a string requires a lookup into this table. For large files (100k+ rows with string data), this becomes a serial bottleneck:
- SheetJS: sequential lookup, no parallelism
- Calamine: string deref per cell is a significant portion of the 2-3x overhead above baseline XML parsing

**5. ZIP Decompression as I/O Bottleneck**

XLSX files are ZIP archives containing XML files. Decompression is CPU-bound and typically single-threaded:
- SheetJS: entire ZIP must decompress into memory before any parsing begins
- ExcelJS: streaming via `node-unzipper` but drops worksheets on files with >100 sheets
- Calamine: worksheets are independent ZIP entries (theoretically parallelizable) but Rust's borrow checker complicates multi-handle access

**6. XML Parsing Overhead**

XLSX uses verbose XML. Parsing dominates I/O time:
- Calamine benchmark: `quick_xml` baseline is 6s for a 1.3GB file; Calamine adds 2-3x overhead on top
- SAX/event-based parsing (quick_xml, ExcelJS streaming) is 5-10x more memory-efficient than DOM-based (POI XSSF)
- Type coercion during parsing (XML text → typed values) is surprisingly expensive

**7. Volatile Function Cascades**

A single NOW() in a workbook marks itself and all dependents as dirty on every recalculation. Chains of volatile functions create exponential dirty propagation. Excel's documentation specifically warns about this.

### Bottleneck Severity by Workload

| Workload | Primary Bottleneck | Secondary |
|----------|-------------------|-----------|
| **Read large XLSX** | ZIP decompression + XML parsing | Shared string table |
| **Write large XLSX** | XML serialization + ZIP compression | Style deduplication |
| **Recalc formula-heavy** | Per-cell eval overhead | Dependency graph traversal |
| **Recalc with ranges** | O(n²) dependency edges | Range aggregation |
| **Interactive editing** | Dirty propagation latency | Re-rendering |
| **Large string-heavy sheets** | String allocation + GC pressure | Shared string lookup |

---

## 17. Cutting-Edge Techniques for Spreadsheet Computation

### Parallel Evaluation: Puncalc (Academic Research)

Puncalc (Journal of Supercomputing, 2019) demonstrates automatic task-based parallel evaluation of spreadsheets:

- Treats spreadsheets as **declarative, purely functional programs** where cell immutability guarantees data-race freedom
- **Topology-agnostic algorithm** — doesn't require a pre-computed topological sort; instead uses speculative evaluation with dynamic dependency discovery
- Achieves **up to 16x speedup on 48 cores** on both synthetic and real-world spreadsheets
- Key insight: exploits both **dataflow parallelism** (independent formula chains) and **data-level parallelism** (SIMD within single formulas)

**Speculative reevaluation:** When a cell's dependencies aren't fully computed, Puncalc speculatively evaluates it anyway. If the speculated dependencies turn out correct, the result is kept; otherwise, the cell is re-evaluated. This avoids the serialization bottleneck of strict topological ordering.

Relevance: Ascend could adopt speculative evaluation for the Web Worker model — start evaluating cells before all dependencies are confirmed, re-evaluate the few that were wrong. For most real workbooks, speculation accuracy would be >95%.

### Asynchronous Computation: DataSpread (Berkeley/UIUC)

DataSpread (SIGMOD '19) tackles the UI-blocking problem for large spreadsheets:

- Both dependency identification and optimal computation scheduling are **NP-Hard** for general spreadsheets
- Solution: **lossy compressed dependency tables** that enable bounded-time dependency identification at the cost of occasionally recomputing cells that didn't need it
- Formulas compute in background with **visual indicators** for pending cells
- Users can continue editing while computation proceeds — the spreadsheet is never "frozen"

This represents the frontier of academic spreadsheet research. The key tradeoff: accept occasional unnecessary recomputation to guarantee responsive UI.

### WasmGC: The Future of Browser Computation

Google Sheets' migration from JS to WasmGC achieved 2x faster calculations (4x improvement from initial WasmGC prototype). Key implications:

- **Garbage-collected languages can run at near-native speed in the browser** via WasmGC
- WasmGC is to Java/Kotlin/Dart what traditional WASM is to C++/Rust
- Future: **shared-memory multithreading** for WasmGC is in development (Chrome status), which would enable Excel-style MTR in the browser
- Ascend implication: if Ascend ever considers a non-JS core, compiling Java/Kotlin to WasmGC is now a viable path alongside Rust→WASM

### GPU Acceleration: Selective, Not Universal

Research shows GPU acceleration is **highly selective** for spreadsheet workloads:

| Operation Type | GPU Benefit | Reason |
|---------------|-------------|--------|
| Matrix multiplication (MMULT) | 2-10x for 500×500+ matrices | Massively parallel, regular access patterns |
| Element-wise arithmetic | **Negative** (14-55ms overhead) | GPU kernel launch latency dominates |
| Large columnar aggregation | 2-4x with WebGPU | Amortizes launch cost over large data |
| String operations | None | Strings are irregular, variable-length |

The Trueno library (Rust) found that GPU is beneficial only for matrix operations above a size threshold. Element-wise operations (the vast majority of spreadsheet formulas) are faster on CPU. **Don't GPU-accelerate spreadsheets broadly — only specific operations on large datasets.**

WebGPU in browsers is now broadly available (Chrome, Edge, Firefox behind flag) and could be used for MMULT on large numeric arrays, but the 14-55ms kernel launch overhead means it's only worthwhile for matrices larger than ~500×500.

### Compiled Formula Execution: The Performance Hierarchy

From slowest to fastest, with real-world examples:

| Approach | Overhead/Eval | Real Example | Perf vs Tree-Walk |
|----------|--------------|--------------|-------------------|
| **Tree-walking interpreter** | ~100-200ns/cell | POI, IronCalc, HyperFormula | 1x (baseline) |
| **Bytecode/stack machine** | ~30-50ns/cell | Excel's internal Ptg evaluator | ~3-4x |
| **Closure-based codegen** | ~15-30ns/cell | Ascend codegen.ts, GRID | ~5-8x |
| **Native JIT** | ~5-10ns/cell | Abacus (JVM bytecode) | ~15-20x |
| **Compiled + SIMD** | ~1-2ns/cell | Throbol (C++ with AVX) | ~100x |

The gap between tree-walking and compiled+SIMD is **two orders of magnitude**. Ascend's closure-based codegen is in the middle of this spectrum — good, but with significant room to improve via WASM compilation and SIMD for numeric ranges.

### Apache Arrow for Spreadsheet Storage

Formualizer's use of Arrow-backed columnar storage is a novel approach:

- **Columnar layout** — all values in a column stored contiguously, ideal for SUM/AVERAGE/COUNTIF over columns
- **O(1) random access** via fixed-width buffers — no tree traversal or hash lookup
- **64-byte aligned** — hardware SIMD and prefetch friendly
- **Zero-copy IPC** — share data between threads/processes/WASM without serialization
- **Dictionary encoding** — efficient storage for columns with repeated string values (categories, labels)

The tradeoff: Arrow is optimized for analytical reads, not mutations. Spreadsheets are inherently mutation-heavy (cell edits). Formualizer handles this with "spill overlays" — mutations go to an overlay that's periodically compacted into the Arrow buffers.

This is analogous to LSM-tree write patterns in databases: buffer writes in a fast mutable layer, periodically merge into the immutable columnar store.

### Incremental Computation Frameworks

Two Rust frameworks directly applicable to spreadsheet recalculation:

**Salsa** (used in rust-analyzer, 2.8k stars):
- Programs defined as queries (pure functions from keys to values)
- Automatic memoization and incremental recomputation
- When an input changes, only transitively dependent queries recompute
- This is *exactly* the spreadsheet recalculation problem, formalized

**Adapton** (academic, general-purpose):
- Demanded computation graph (DCG) — tracks dependencies automatically
- Propagates changes minimally through the graph
- More theoretical, less production-tested than Salsa

Both implement the theoretical optimum: minimal recomputation. The question for Ascend is whether the framework overhead (query registration, cache management) is worth it compared to a hand-tuned dirty-flag propagation system.

### Benchmarking Spreadsheet Systems (SIGMOD '20)

The DataSpread team's systematic benchmark revealed that current spreadsheet systems **vastly underperform** database-backed alternatives on large datasets:

- Importing 100,000 rows in Excel takes over 10 minutes
- Spreadsheet systems lack fundamental database optimizations: indexing, intelligent data layout, shared computation
- Column operations (SUM, AVERAGE) that should be O(1) with columnar storage are O(n) in every tested system

This suggests the biggest performance opportunity isn't in formula evaluation speed but in **storage layout**: columnar/indexed storage for the common case of column-wide aggregation.

---

## 18. Comparative Summary

| Engine | Cell Storage | Formula Eval | Memory Model | Key Innovation |
|--------|-------------|--------------|--------------|----------------|
| **SheetJS** | Object props or Array[][] | None | Full in-memory | Dense mode simplicity |
| **ExcelJS** | Class instances (Row→Cell) | None | Full in-memory | Streaming async iterators |
| **OpenPyXL** | Python objects | None | Read-only/write-only modes | Generator-based bounded iteration |
| **Apache POI** | XML DOM (XSSF) / byte arrays (HSSF) | Tree-walk, 140 functions, cached | Sliding window (SXSSF) | Evaluation cache + tracker |
| **Calamine** | Flat Vec + Data enum | None | Contiguous row-major | Minimal overhead, SSO exploration |
| **IronCalc** | BTreeMap<row, BTreeMap<col, Cell>> | Tree-walk, 192 functions | Full in-memory | WASM-first, multi-language bindings |
| **LibreOffice** | multi_type_vector (contiguous typed blocks) | Threaded formula groups | Broadcaster-listener | Typed block storage, formula group parallelism |
| **HyperFormula** | Internal graph nodes | Tree-walk, 400+ functions | Full in-memory | Range decomposition, AST reuse, lazy CRUD |
| **GRID** | Internal (undisclosed) | AST evaluation | Full in-memory | Shared eval context, real-workbook benchmarks |
| **Univer** | Plugin-managed, Canvas | Tree-walk, NumPy-style arrays | Web Worker isolation | ArrayValueObject, Lambda, feature plugins |
| **Formualizer** | Arrow columnar + overlays | Compiled, 320+ funcs, Rayon parallel | Arrow zero-copy | CSR dep-graph, deterministic mode |
| **Throbol** | Scratchpad (flat buffer) | Compiled closures + SIMD | Pre-allocated pad | 11.5B opcodes/s, SIMD batching |
| **Google Sheets** | Server-managed | WasmGC (Java to WASM) | Worker-isolated | 2x faster than JS via WasmGC |
| **Excel** | Native C++ (proprietary) | MTR, 1024 threads, 500+ funcs | Partitioned memory | Smart recalc, chain reordering |

---

## 19. Actionable Recommendations for Ascend

### High-Impact, Low-Effort

1. **Shared mutable EvalContext** — Stop creating evaluation context objects per cell. Single frozen context with mutable cell/ref fields. GRID proved ~10% recalc speedup. *Estimated effort: 1-2 days.*

2. **Lazy context properties** — Use getters for sheet resolution, workbook mode, table context. Only pay for what the formula uses. *Pairs with #1.*

3. **Range node decomposition in dep-graph** — When SUM(A1:A100) and SUM(A1:A101) both exist, represent A1:A101 as A1:A100 + A101. Prevents O(n²) edge explosion. *Estimated effort: 2-3 days.*

### Medium-Impact, Medium-Effort

4. **Expand codegen coverage** — Add SUM, AVERAGE, COUNT, MIN, MAX, VLOOKUP to the codegen path. These are the most common formulas and benefit most from eliminating tree-walk overhead. *Estimated effort: 1 week.*

5. **Formula group detection and batch evaluation** — Detect columns where every cell has the same formula structure (just offset references). Generate one function, call it N times with row offsets. Combine with the dep-graph for batch dirty-marking. *Estimated effort: 1-2 weeks.*

6. **Lazy structural transformations** (HyperFormula pattern) — On row/column insert/delete, don't immediately update all formula references. Mark formulas as "needs reference adjustment" and defer until recalc. *Estimated effort: 1 week.*

### High-Impact, High-Effort

7. **Contiguous typed storage within chunks** — Evolve SparseGrid's dense chunks toward LibreOffice's multi_type_vector concept: within a chunk, store runs of same-typed values in contiguous arrays. Enables SIMD aggregation and better cache utilization. *Estimated effort: 2-3 weeks.*

8. **WASM numeric kernel** — Compile Rust SIMD aggregation functions (sum, min, max, countif over f64 arrays) to WASM. Feed them contiguous numeric ranges from the grid. 2-4x speedup on range aggregations. *Estimated effort: 2-3 weeks.*

9. **Async/threaded recalculation** — Use Web Workers or `SharedArrayBuffer` to parallelize independent branches of the calculation chain (following Excel's MTR model). Requires thread-safe dep-graph and partitioned cell access. *Estimated effort: 1-2 months.*

### Research/Long-term

10. **Salsa-style incremental computation** — Replace the current dep-graph + dirty-flag + topological sort with a query-based incremental framework. Each cell value becomes a memoized query. Changes propagate minimally. This is architecturally ambitious but represents the theoretical optimum. *Estimated effort: research project.*

---

## 20. Google Sheets WasmGC Migration — The Landmark Case Study

**Source:** [web.dev case study (2024)](https://web.dev/case-studies/google-sheets-wasmgc)

### The Problem

Google Sheets' calculation engine was originally written in Java (2006). It moved client-side in 2013 via GWT→J2CL transpilation to JavaScript. Performance testing revealed the JavaScript version was **>3x slower than original Java**.

### Architecture

The calc engine runs in a **Web Worker**, communicating with the main thread via `MessageChannel`. This isolates calculation from rendering — a clean separation Ascend already follows via its pure-function engine design.

### WasmGC: What It Is

WasmGC (WebAssembly Garbage Collection) extends the Wasm spec with primitives for compiling garbage-collected languages (Java, Kotlin, Dart) to Wasm. Unlike standard Wasm (for C/C++/Rust), it handles reference types and GC natively in the VM rather than bringing a custom allocator.

### Performance Results

| Version | Relative Performance |
|---------|---------------------|
| Original Java (server) | 1.0x baseline |
| JavaScript (J2CL) | 0.33x (3x slower) |
| WasmGC (initial, unoptimized) | 0.17x (6x slower) |
| WasmGC (optimized, final) | 0.67x (1.5x slower than Java, **2x faster than JS**) |

4x improvement from initial WasmGC to final, via targeted optimizations.

### Key Optimizations They Made

1. **Speculative inlining + devirtualization** — Implementing what JVMs do automatically (inline caching for virtual dispatch). Yielded **~40% speedup** alone.

2. **Browser API delegation for strings and regex** — Compiling `re2j` (a Java regex engine) to WasmGC was 100x slower than calling the browser's native `RegExp` API. Lesson: don't reimplement what the host does natively.

3. **Platform-agnostic data structures** — JavaScript's auto-promotion of sparse arrays to maps was relied on in the J2CL code. This "optimization" became a deoptimization in WasmGC. They rewrote to explicit data structures.

### Relevance to Ascend

Google chose WasmGC because their engine was already Java. Ascend's engine is TypeScript — we don't have a pre-existing compiled language to port. But two findings directly apply:

- **Hot numeric kernels benefit from compilation.** Google got 2x over JS for the calculation worker. Ascend's existing WASM range ops (`wasm-range.ts`) follow the same pattern — compile the tight numeric loops, keep orchestration in TS.
- **Don't fight the host platform.** String operations, regex, JSON parsing — these are faster via JS APIs than compiled alternatives. Keep formula logic that touches strings in TS.

---

## 21. Formualizer — Arrow-Native Rust Engine (2025-2026)

**Language:** Rust + WASM | **Stars:** New, actively developed | **Source:** [formualizer.dev](https://www.formualizer.dev/), [docs.rs/formualizer-eval](https://docs.rs/formualizer-eval)

### Architecture

Modular pipeline: `formualizer-parse` → `formualizer-eval` → `formualizer-workbook`. Built on Apache Arrow for columnar data.

### Key Technical Decisions

1. **Compressed Sparse Row (CSR) for dependency edges.** The dep graph stores edges in CSR format — a flat array of destination indices plus a pointer array marking where each node's edges begin. This is the most cache-friendly graph representation for iteration-heavy algorithms like topological sort.

2. **Parallel evaluation via Rayon.** Independent subgraphs evaluate on separate threads using Rust's Rayon work-stealing scheduler. The scheduler automatically balances work across cores.

3. **Deterministic evaluation mode.** Injected clock, timezone, and RNG seeds for reproducible results. Critical for testing and for agent workflows where reproducibility matters.

4. **320+ Excel functions.** More than HyperFormula (400+ claimed but many are stubs), competitive with commercial engines.

### CSR Format Deep Dive

Traditional adjacency list:
```
Node 0 → [1, 3, 5]
Node 1 → [2]
Node 2 → [4]
```

CSR representation:
```
edges:    [1, 3, 5, 2, 4]     // all edges, flattened
offsets:  [0, 3, 4, 5]         // edges[offsets[i]..offsets[i+1]] = neighbors of node i
```

Benefits: contiguous memory, no per-node allocation, SIMD-friendly iteration, predictable cache access. For a dep graph with 100k nodes and 500k edges, CSR uses ~2.4 MB vs ~12 MB for a `Map<number, number[]>` with per-array overhead.

### Worth Stealing

- **CSR for the dep graph.** Ascend's dep graph uses Maps. Switching to CSR for the forward-edge representation (used during topological sort and dirty propagation) would improve cache locality for recalculation walks. The adjacency data is rebuilt on structural changes anyway.
- **Deterministic mode with injected time/RNG.** Ascend's agent-native design should guarantee reproducible evaluation. This is table-stakes for testability.

---

## 22. Equals Engineering — Rectangle Consolidation for Metadata

**Source:** [Building a High Performance Spreadsheet (2024)](https://engineering.equals.com/p/building-a-high-performance-spreadsheet)

### The Problem

The Equals team discovered that the biggest perf bottleneck in their browser spreadsheet wasn't formula calculation — it was **Formats** (runtime type info, background colors, number formats). Their Head of Finance's workbooks became unusable.

### Evolution of Failed Approaches

| Approach | Problem |
|----------|---------|
| Per-cell metadata objects | Write fan-out: selecting an entire column creates millions of updates |
| Per-row/column storage | Only helps if the write covers an exact row/column |
| Append-log (last-write-wins) | Reads degrade over time as the log grows |

### Solution: R-Tree + Greedy Rectangle Consolidation

Three components:

1. **Core storage:** Plain JS objects in Maps keyed by numeric IDs. V8 optimizes statically-shaped plain objects extremely well.

2. **Read index:** An **R-tree** spatial index ([rbush](https://github.com/mourner/rbush)) for fast "give me all formats intersecting this range" queries. R-trees provide O(log n) range queries over 2D data.

3. **Write consolidation:** A greedy algorithm that repeatedly merges adjacent rectangles sharing an edge. Not guaranteed optimal (a perfect packing is O(n^{3/2})), but:
   - O(n) worst case
   - Stable — small edits produce small diffs
   - Zero pathological cases observed in production

### Relevance to Ascend

Ascend's `StyleRegistry` already handles style interning (avoiding ExcelJS's catastrophic per-cell style objects). But for metadata that covers ranges (number formats, conditional formatting, data validation), an R-tree index could dramatically improve lookup performance. The greedy rectangle consolidation is a practical alternative to optimal rectangle partitioning.

---

## 23. Salsa — Deep Dive on the Red-Green Algorithm

**Source:** [salsa-rs.github.io/salsa/reference/algorithm](https://salsa-rs.github.io/salsa/reference/algorithm.html), [Salsa Algorithm Explained (Medium)](https://medium.com/@eliah.lakhin/salsa-algorithm-explained-c5d6df1dd291)

### The Algorithm

Salsa maintains a DAG of pure functions with automatic dependency tracking. Three core mechanisms:

**1. Revision Tracking**

The database maintains a global revision counter, incremented on every input `set`. Each memoized result stores:
- The return value
- The revision it was computed in
- A list of dependencies with the revision each dependency last changed

On re-invocation, if the current database revision equals the result's revision, return immediately. Otherwise, walk dependencies to check if any actually changed.

**2. Backdating**

When a tracked function re-executes and produces the **same output** as before, Salsa "backdates" the result: it marks the output as unchanged even though inputs changed. This means downstream consumers won't re-execute.

Example for spreadsheets: `=LEN(A1)` depends on A1. If A1 changes from "hello" to "world", LEN(A1) still returns 5. Backdating prevents all cells depending on LEN(A1) from recalculating.

**3. Durability**

An optimization for multi-frequency inputs. High-durability inputs (e.g., imported reference data, static tables) rarely change. Low-durability inputs (e.g., the cell being edited) change constantly.

The database tracks the last-changed revision **per durability level**. When checking a memoized result that only depends on high-durability inputs, if no high-durability input has changed since the result was computed, skip the dependency walk entirely.

For spreadsheets: imported data tables → high durability. User-edited cells → low durability. Most formulas referencing imported data skip recalculation when the user edits unrelated cells.

### Mapping to Spreadsheet Recalculation

| Salsa Concept | Spreadsheet Equivalent |
|---------------|----------------------|
| Input | Cell with a literal value |
| Tracked function | Cell with a formula |
| Revision | Edit generation counter |
| Dependency | Cell reference in a formula |
| Backdating | "Output didn't change" short-circuit |
| Durability | Static data vs. user-edited cells |

### Implementation Sketch for Ascend

```typescript
interface MemoEntry {
  value: CellValue
  verifiedAt: number       // revision when last verified valid
  changedAt: number        // revision when value last actually changed  
  deps: { cellId: string; changedAt: number }[]
  durability: 'high' | 'low'
}

function queryCell(cellId: string, revision: number): CellValue {
  const memo = cache.get(cellId)
  if (memo && memo.verifiedAt === revision) return memo.value
  
  // Durability fast-path
  if (memo && memo.durability === 'high' && lastHighDurabilityChange < memo.verifiedAt) {
    memo.verifiedAt = revision
    return memo.value
  }
  
  // Check if any dep actually changed
  if (memo && memo.deps.every(d => queryCell(d.cellId, revision) === /* unchanged */)) {
    memo.verifiedAt = revision
    return memo.value  // deps unchanged, reuse result
  }
  
  // Re-evaluate
  const newValue = evaluate(cellId)
  const changed = !deepEqual(newValue, memo?.value)
  cache.set(cellId, {
    value: newValue,
    verifiedAt: revision,
    changedAt: changed ? revision : (memo?.changedAt ?? revision),
    deps: currentDeps,
    durability: inferDurability(cellId),
  })
  return newValue
}
```

### Why This Matters

Ascend's current recalculation does topological sort + dirty propagation. Salsa's approach is **demand-driven** — it only verifies cells that are actually observed (visible in the viewport, queried by an agent, or needed by a dependent). Combined with backdating, this can dramatically reduce recalculation work when changes propagate through intermediate formulas that produce the same result.

---

## 24. Jane Street's Incremental — Production Self-Adjusting Computation

**Source:** [blog.janestreet.com/introducing-incremental](https://blog.janestreet.com/introducing-incremental/), [Seven Implementations of Incremental](https://www.janestreet.com/tech-talks/seven-implementations-of-incremental/)

### Key Design Differences from Salsa

| Property | Salsa | Incremental |
|----------|-------|-------------|
| Evaluation trigger | On demand (call the function) | Explicit `stabilize()` call |
| Graph structure | Static (compile-time) | Dynamic (can change at runtime via `bind`) |
| Cutoff mechanism | Backdating (same output → unchanged) | Physical equality cutoff (configurable) |
| Primary use | Compilers (rust-analyzer) | GUIs, trading systems |

### Physical Equality Cutoff

Incremental's default cutoff: if the new value is **physically equal** (===) to the old value, don't propagate. This is configurable — you can define custom cutoffs like "changed by less than epsilon" for floating-point values.

For spreadsheets, a custom cutoff could prevent cascading recalculation when rounding differences are irrelevant.

### Dynamic Graph Structure

Unlike spreadsheets (where the dependency graph is determined by formula content), Incremental supports `bind` — a node that can change its structure based on its inputs. Example: `IF(A1 > 0, B1, C1)` depends on B1 OR C1, not both, depending on A1's value.

Ascend could implement conditional dependency tracking: when evaluating `IF`, only register a dependency on the taken branch. This reduces false dirty propagation.

### Relevance to Ascend

- **Explicit stabilize model** matches batch-edit patterns. An agent makes 50 cell edits, then asks for results. Only one stabilization pass needed, not 50.
- **Conditional dependency tracking** for IF/SWITCH/IFS could reduce the dirty set substantially. Current approach marks both branches as dependencies.

---

## 25. Incremental Topological Sort — State of the Art (2024)

### McCauley et al. (ICML 2024) — Ordering with Predictions

**Source:** [proceedings.mlr.press/v235/mccauley24a](https://proceedings.mlr.press/v235/mccauley24a.html)

The paper introduces incremental topological ordering that uses **learned predictions** about the graph structure. When edges arrive one-by-one (cells added, formulas changed), the algorithm:

1. Maintains a topological ordering consistent with all observed edges
2. Uses predictions (trained on historical graph patterns) to pre-position new nodes
3. Guarantees correctness regardless of prediction quality (robustness)
4. Achieves optimal runtime when predictions are accurate

For Ascend: spreadsheet dependency graphs have highly predictable structure (most formulas reference nearby cells, column patterns repeat). A simple heuristic — "new formulas in column C probably depend on columns A-B in the same row" — could serve as the prediction, avoiding expensive reordering.

### Parallel Topological Sort

**Source:** [CMU CS project (2024)](https://github.com/codeplay0314/parallel-topological-sorting)

Three approaches benchmarked:

| Approach | Best For | Speedup |
|----------|----------|---------|
| Serial Kahn's algorithm | Baseline | 1x |
| CUDA GPU parallel | Very large graphs (1M+ nodes) | 5-50x |
| Boost Graph Library parallel | Moderate graphs, CPU-only | 2-4x |

For Ascend's scale (typical workbooks: 1k-100k formula cells), the serial Kahn's algorithm with incremental updates is likely optimal. GPU parallelism only pays off at scales beyond typical spreadsheets.

### Practical Incremental Update Strategy

Instead of re-sorting the entire graph on each edit:

```
On cell edit(cellId):
  1. Identify affected subgraph (cells reachable from cellId in forward direction)
  2. Only re-sort the affected subgraph
  3. Splice the re-sorted subgraph back into the global ordering
```

For local edits (which dominate), the affected subgraph is typically <1% of the total graph.

---

## 26. Memory-Efficient Data Structures for JS/TS Engines

### Generational Arena Allocator

**Source:** [generational-arena (JS)](https://github.com/richardanaya/generational-arena)

Arena allocators pre-allocate a contiguous block and hand out indices instead of heap pointers. The "generational" variant adds a generation counter to detect use-after-free.

```typescript
// Conceptual arena for cell values
class CellArena {
  private values: Float64Array    // numeric values
  private generations: Uint32Array // generation per slot
  private freeList: number[]       // available slots
  
  alloc(): { index: number; generation: number } { ... }
  get(handle: { index: number; generation: number }): number { ... }
  free(handle: { index: number; generation: number }): void { ... }
}
```

Benefits for spreadsheets:
- No GC pressure from cell value allocations
- Cache-friendly contiguous storage
- O(1) alloc/free via free list
- Generation counter prevents stale-reference bugs

### HAMT (Hash Array Mapped Trie) for Persistent Data Structures

**Source:** [hamt (JS)](https://github.com/mattbierner/hamt), [immutable-collections (TS)](https://immutable-collections.seedtactics.com/)

HAMTs enable structural sharing: updating one cell in a million-cell workbook only allocates O(log n) new nodes, reusing the rest. This is how Clojure's persistent vectors and Immutable.js work.

For undo/redo and collaborative editing, HAMTs provide:
- O(log₃₂ n) lookup, insert, delete (practically O(1) for n < 1B)
- Efficient diffing between versions (shared prefixes skip comparison)
- Memory-efficient snapshots (each version shares structure with previous)

### Compressed Sparse Row for the Dependency Graph

**Source:** [csr-matrix (JS)](https://github.com/mikolalysenko/csr-matrix), [@thi.ng/sparse (TS)](https://docs.thi.ng/umbrella/sparse/classes/CSR.html)

The `@thi.ng/sparse` library provides a TypeScript CSR implementation with methods for matrix operations. For Ascend's dep graph:

- Store all dependency edges in a single `Int32Array`
- Store offsets in a second `Int32Array` 
- Total memory: ~8 bytes per edge (vs ~40+ bytes for Map entries)
- Iteration over a node's dependents is a tight loop over contiguous memory

Rebuild cost is O(V + E) when the graph structure changes, but reads during recalculation (the hot path) become cache-friendly array scans.

---

## 27. Modern I/O Techniques

### Zero-Copy Parsing in JavaScript

**Source:** [foxglove/mcap PR #1185](https://github.com/foxglove/mcap/pull/1185), [zero-copy optimization](https://haikel-fazzani.deno.dev/blog/zero-copy-technique)

Key techniques:

**1. `Buffer.subarray()` over `Buffer.slice()`**

`subarray` creates a view sharing the same memory. `slice` copies. For XLSX parsing where we're extracting thousands of XML fragments from a decompressed buffer, this is the difference between O(1) and O(n) per extraction.

**2. TypedArray construction optimization**

```typescript
// Slow (copy + copy):
new Uint8Array(buffer.slice(start, end))

// Fast (view + copy, avoids intermediate allocation):
new Uint8Array(buffer, start, length).slice()
```

Benchmark: 1.29 ops/s → 3.04 ops/s (2.4x improvement) for stream parsing.

**3. Transferable Objects for Workers**

`ArrayBuffer` can be transferred (not cloned) to a Web Worker via `postMessage(data, [data.buffer])`. The transfer is O(1) regardless of buffer size — the ownership moves, no copy occurs.

**4. SharedArrayBuffer for parallel access**

Multiple threads read/write the same memory region. Combined with `Atomics` for synchronization. Enables a shared cell-value buffer accessible from both the main thread and calculation workers without serialization.

### Streaming XML Parsers

| Parser | Style | Features |
|--------|-------|----------|
| [StAX-XML](https://clickin.github.io/stax-xml/) | Pull (StAX) | Bun/Deno/Node/browser, streaming |
| [xlsx-stream-reader](https://github.com/DaSpawn/xlsx-stream-reader) | SAX events | Non-blocking, memory-efficient |
| [xlstream](https://www.npmjs.com/package/xlstream) | Transform stream | Pausable, merged cells, formatting |

For Ascend's `io-xlsx` reader: a pull parser (StAX) gives more control than SAX events. The reader can pull exactly the elements it needs (cells, shared strings, styles) and skip everything else (drawings, charts, print settings) without registering callbacks.

### xlsx-fire: Rust+WASM XLSX Reader

A new entrant: Rust core compiled to WASM, providing "blazing-fast, memory-safe async streaming" for large XLSX files. If Ascend wanted to accelerate XLSX reading beyond what pure JS can achieve, a Rust→WASM reader for the decompression + XML parse hot path is a proven approach.

---

## 28. Bun-Specific Optimizations

### Bun FFI (`bun:ffi`)

**Source:** [bun.sh/reference/bun:ffi](https://bun.sh/reference/bun:ffi)

Bun's FFI uses **embedded TinyCC** to JIT-compile C bindings for type conversion, achieving 2-6x faster native calls than Node.js FFI (Node-API).

Practical constraints:
- Only supports C ABI (numbers and pointers)
- Complex data requires manual pointer handling via TypedArrays
- TypedArrays are automatically converted to pointers
- Still experimental — Node-API modules are more stable for production

**Opportunity for Ascend:** A small C/Zig library for hot-path operations (SUM over f64 arrays, string hashing for shared string tables, CRC32 for ZIP) callable via `bun:ffi` could outperform both JS and WASM, since FFI avoids WASM's function call overhead.

### Bun's Native SQLite

**Source:** [bun.sh/docs/runtime/sqlite](https://bun.sh/docs/runtime/sqlite)

Built-in `bun:sqlite` is synchronous and claims 3-6x faster reads than `better-sqlite3`. However:
- Write performance has known issues (~957 writes/s on SSD vs 51k/s for better-sqlite3 in early benchmarks)
- Complex queries (joins, aggregations) may be slower than better-sqlite3
- In-memory performance is more competitive (67k writes/s)

**Potential for Ascend:** SQLite as a workbook persistence format for large workbooks. Store cell data in a table with (sheet, row, col, value, formula) columns, styles in a separate table. SQLite's B-tree storage handles sparse data naturally, and its WAL mode supports concurrent reads during writes. But the write perf issues need resolution first.

### JavaScriptCore Optimization Patterns

Bun runs on JSC (Safari's engine), not V8. Key differences:

1. **Different JIT tiers:** JSC uses LLInt → Baseline JIT → DFG → FTL (via B3/LLVM-derived backend). V8 uses Ignition → Sparkplug → Maglev → TurboFan.

2. **Object model:** JSC uses "butterfly" storage (contiguous property + element storage), V8 uses hidden classes/maps. Both benefit from consistent object shapes.

3. **Optimization patterns that work across both engines:**
   - Initialize all properties in constructors
   - Add properties in consistent order
   - Avoid `delete` on objects (use `null` assignment)
   - Keep hot-path functions monomorphic (always receive same types)
   - Prefer `for` loops over `forEach`/iterators in hot paths

---

## 29. JS Engine Optimization — Hidden Classes and Inline Caches

**Source:** [v8.dev/docs/hidden-classes](https://v8.dev/docs/hidden-classes), [mathiasbynens.be/notes/shapes-ics](https://mathiasbynens.be/notes/shapes-ics)

### Why This Matters for Ascend

Every property access in the formula evaluator and dep graph goes through the engine's property lookup mechanism. The difference between monomorphic (optimized) and megamorphic (deoptimized) access can be 10-100x.

### Hidden Classes (Shapes/Maps)

When you create `{ kind: 'number', value: 42 }`, V8/JSC creates a hidden class describing the layout: `kind` at offset 0, `value` at offset 8. Subsequent objects with the same property sequence reuse this hidden class.

**Critical for CellValue:** Ascend's `CellValue` type is `{ kind: 'number' | 'string' | 'boolean' | 'error' | 'empty', value: ... }`. If all CellValue objects are created with properties in the same order, they share hidden classes and get fast access. If some are created as `{ kind, value }` and others as `{ value, kind }`, each gets a different hidden class.

### Inline Cache States

| State | # of shapes seen | Performance |
|-------|------------------|-------------|
| Monomorphic | 1 | Fastest (direct offset access) |
| Polymorphic | 2-4 | Fast (small lookup table) |
| Megamorphic | 5+ | Slow (hash table fallback) |

### Actionable Rules for Ascend's Hot Paths

1. **Factory functions for CellValue** — Always create via `numberValue()`, `stringValue()`, etc. to guarantee consistent property order.
2. **Avoid optional properties** — `{ kind: 'number', value: 42, error: undefined }` and `{ kind: 'number', value: 42 }` have different hidden classes. If some cells have `formula` and others don't, always include the property (set to null).
3. **Monomorphic eval dispatch** — The compiled-eval VM switch statement is fine (integer switch is well-optimized). But any function that receives `CellValue` should only see CellValue objects from the factory functions, not ad-hoc objects.
4. **Avoid `delete`** — Never `delete cell.formula`. Set `cell.formula = null`.

---

## 30. Agent-Native Spreadsheet Design

**Source:** [MCP Best Practices](https://mcp-best-practice.github.io/mcp-best-practice/), [7 Principles for AI Agent Tool Design](https://dev.to/alexchen31337/7-principles-for-ai-agent-tool-design-from-claude-code-real-world-systems-3dcd), [Designing APIs for AI Agents First](https://dev.to/codimow/designing-apis-for-ai-agents-first-what-the-model-context-protocol-taught-us-2j46)

### Core Principles

**1. Atomicity Over Composability**

Each tool does one thing and returns one result. Agents cannot reliably compose multi-step API sequences — they hallucinate intermediate steps, skip error handling, or misinterpret relationships.

Bad:
```
get_cell(A1) → value
set_cell(A1, value + 1) → ok
```

Good:
```
increment_cell(A1) → { old: 5, new: 6 }
```

For Ascend's MCP server: operations like "sort this range by column B" should be a single tool call, not "read range → sort in memory → write range."

**2. Progressive Disclosure**

Don't dump the entire workbook schema into context. Provide search/discovery tools:
- `describe_workbook()` → sheet names, dimensions, named ranges
- `search_cells(query)` → cells matching a pattern
- `get_range(ref)` → values in a specific range

The agent pulls what it needs. Prevents "context rot" where irrelevant information degrades decision quality.

**3. Structured Output Over Prompt Instructions**

Use tool parameter schemas (Zod with `.describe()`) to enforce structure rather than asking the model to "format your output as JSON." Schema-enforced outputs eliminate parsing errors.

**4. Flat Arguments**

```typescript
// Bad: nested config objects
{ range: "A1:B10", options: { sort: { column: "B", direction: "asc" } } }

// Good: flat primitives
{ range: "A1:B10", sort_column: "B", sort_direction: "asc" }
```

**5. Namespace Tools at Scale**

For >10-20 tools, use forward-slash namespacing:
- `cells/read`, `cells/write`, `cells/format`
- `formulas/evaluate`, `formulas/trace`
- `sheets/create`, `sheets/delete`

### MCP-Specific Patterns

| Pattern | Description |
|---------|-------------|
| **Outcomes over operations** | `track_latest_order(email)` instead of three separate API calls |
| **Stateless by default** | Each tool call is independent; no session state between calls |
| **Contracts first** | Strict input/output schemas, explicit side effects, documented errors |
| **~20 tool limit** | Beyond this, agent decision quality degrades |

### Ascend-Specific Recommendations

1. **Batch operations tool** — `apply_operations([{set: "A1", value: 42}, {set: "B1", formula: "=A1*2"}])` — reduces round-trips for multi-cell edits.
2. **Explain tool** — `explain_formula("=VLOOKUP(A1,Data!A:C,3,FALSE)")` → natural language explanation of what the formula does, its dependencies, potential errors.
3. **Diff tool** — `diff_workbook(before_snapshot, after_snapshot)` → structured diff of what changed, enabling agents to verify their edits.
4. **Validate tool** — `validate_workbook()` → structural issues, broken references, circular deps, type mismatches. Gives agents self-repair capability.

---

## 31. TurboSheets — Competitive Architecture Analysis

**Source:** [How we built a spreadsheet editor 5x faster than Google Sheets](https://turbosheets.notion.site/)

### Performance Claim

~5x faster calculation than Google Sheets (pre-WasmGC migration). Focused on datasets >20k rows where Google Sheets degrades.

### Canvas Rendering Architecture (from industry analysis)

Production browser spreadsheets use **multi-layer HTML5 Canvas**:

| Layer | Content | Why Separate |
|-------|---------|--------------|
| Grid | Cell content, borders | Redraws on scroll/edit |
| Headers | Row/column headers | Fixed position, redraws on resize |
| Selection | Highlight, cursor | Redraws frequently, needs transparency |
| Collaboration | Other users' cursors | Independent update cycle |

Each layer is a separate `<canvas>` element. Only the affected layer redraws on each event. A single spreadsheet has ~17 billion potential cells; rendering is viewport-based with buffer zones (similar to virtual scrolling in list UIs).

### Key Lesson: Metadata Is the Bottleneck

Both Equals and TurboSheets independently discovered that **format/metadata access** — not formula calculation — was their primary performance bottleneck. This suggests Ascend should profile real workbooks and measure time spent in style/format resolution vs formula evaluation.

---

## 32. Persistent Data Structures for Undo/Redo and Collaboration

### Structural Sharing via HAMT

For workbook snapshots (undo/redo, version history, collaborative conflict resolution):

```
Workbook v1: [A1=1, A2=2, A3=3, ..., A1000=1000]
                    ↓ edit A5=99
Workbook v2: [A1=1, A2=2, A3=3, A4=4, A5=99, A6=6, ..., A1000=1000]
```

With structural sharing, v1 and v2 share >99% of their memory. Only the path from A5 to the root is allocated fresh (O(log n) nodes).

Libraries:
- [immutable-collections](https://immutable-collections.seedtactics.com/) — Modern TS, zero deps, tree-shakeable function-based API
- [hamt](https://github.com/mattbierner/hamt) — Minimal HAMT implementation
- [hamt_plus](https://github.com/mattbierner/hamt_plus) — HAMT with transactions and custom key types

### Practical Application for Ascend

Ascend's operation/patch system already provides undo/redo. But for **snapshots** (e.g., "show me what this workbook looked like 10 edits ago"), persistent data structures avoid cloning the entire workbook. A HAMT-based cell store would make snapshot creation O(1) and lookup O(log₃₂ n).

---

## Revised Actionable Recommendations

### Tier 1: Immediate Wins (Days)

| # | Technique | Source | Expected Impact | Effort |
|---|-----------|--------|-----------------|--------|
| 1 | **Backdating optimization** — When a formula re-evaluates to the same value, don't propagate dirty flags downstream | Salsa | 10-30% fewer recalculations for typical edits | 1-2 days |
| 2 | **Durability classification** — Mark imported/static data as high-durability, skip dep-walk when only low-durability inputs change | Salsa | Major speedup for workbooks mixing imported data with user edits | 1 day |
| 3 | **Consistent CellValue factories** — Ensure all CellValue creation goes through factory functions for monomorphic hidden classes | V8/JSC optimization | 5-15% on eval hot path | 0.5 days |

### Tier 2: Medium-Term (Weeks)

| # | Technique | Source | Expected Impact | Effort |
|---|-----------|--------|-----------------|--------|
| 4 | **CSR dep graph edges** — Store forward edges in `Int32Array` with offset index | Formualizer | Better cache locality during recalc walks, ~4x less memory for edges | 1 week |
| 5 | **Demand-driven recalculation** — Only verify cells that are observed (viewport, agent query), not the entire dirty set | Jane Street Incremental | Major speedup for large workbooks where most cells aren't visible | 1-2 weeks |
| 6 | **Conditional dependency tracking** — IF/SWITCH only register dependency on the taken branch | Jane Street Incremental | Reduces false dirty propagation for conditional formulas | 1 week |
| 7 | **Pull-parser for XLSX** — Replace SAX with StAX-style pull parser for selective element reading | StAX-XML | Faster XLSX reads by skipping irrelevant elements | 1 week |
| 8 | **Zero-copy buffer views** — Use `subarray` and transferable objects in XLSX parsing pipeline | Zero-copy techniques | 2x improvement in parse throughput | 3-5 days |

### Tier 3: Strategic (Months)

| # | Technique | Source | Expected Impact | Effort |
|---|-----------|--------|-----------------|--------|
| 9 | **Agent-native MCP redesign** — Atomic operations, progressive disclosure, flat args, batch operations | MCP best practices | Dramatically better agent interaction quality | 2-3 weeks |
| 10 | **Rust→WASM hot kernel** — Dep graph traversal + numeric aggregation in a single WASM module | Google Sheets WasmGC, Formualizer | 2x overall recalc perf for numeric-heavy workbooks | 1-2 months |
| 11 | **R-tree for range metadata** — Spatial index for formats, conditional formatting, data validation | Equals Engineering | O(log n) range metadata lookups vs current approach | 2-3 weeks |
| 12 | **Persistent HAMT cell store** — O(1) snapshots, O(log n) lookups, structural sharing for undo/collab | HAMT research | Efficient version history and collaborative conflict resolution | 1-2 months |

---

## References

- [SheetJS Common Spreadsheet Format](https://docs.sheetjs.com/docs/csf)
- [ExcelJS Model](https://github.com/exceljs/exceljs/blob/master/MODEL.md)
- [OpenPyXL Optimised Modes](https://openpyxl.readthedocs.io/en/stable/optimized.html)
- [Apache POI Formula Evaluation](https://poi.apache.org/components/spreadsheet/eval.html)
- [Calamine Data enum](https://docs.rs/calamine/latest/calamine/enum.Data.html)
- [IronCalc](https://www.ironcalc.com/)
- [mdds::multi_type_vector](https://kohei.us/2012/07/20/mdds-multi_type_vector-explained/)
- [HyperFormula Dependency Graph](https://hyperformula.handsontable.com/guide/dependency-graph.html)
- [GRID Engine Performance](https://alexharri.com/blog/grid-engine-performance)
- [Excel Recalculation (Microsoft)](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-recalculation)
- [Excel MTR (Microsoft)](https://learn.microsoft.com/en-us/office/client-developer/excel/multithreaded-recalculation-in-excel)
- [DataSpread Async Computation (SIGMOD '19)](https://people.eecs.berkeley.edu/~adityagp/papers/dataspread-async.pdf)
- [Benchmarking Spreadsheet Systems (SIGMOD '20)](https://dl.acm.org/doi/10.1145/3318464.3389782)
- [Adapton](https://github.com/Adapton/adapton.rust)
- [Salsa](https://github.com/salsa-rs/salsa)
- [Abacus Formula Compiler](https://formulacompiler.org/doc/design.htm)
- [LibreOffice Calc Architecture](https://niocs.github.io/LOBook/calcarch/overview.html)
- [LibreOffice Threaded Calc](https://www.phoronix.com/news/LibreOffice-Calc-Threading)
- [Google Sheets WasmGC Case Study](https://web.dev/case-studies/google-sheets-wasmgc)
- [Formualizer Eval Engine](https://docs.rs/formualizer-eval)
- [Formualizer Docs](https://www.formualizer.dev/)
- [Equals: Building a High-Performance Spreadsheet](https://engineering.equals.com/p/building-a-high-performance-spreadsheet)
- [Salsa Red-Green Algorithm](https://salsa-rs.github.io/salsa/reference/algorithm.html)
- [Salsa Algorithm Explained](https://medium.com/@eliah.lakhin/salsa-algorithm-explained-c5d6df1dd291)
- [Jane Street: Introducing Incremental](https://blog.janestreet.com/introducing-incremental/)
- [Jane Street: Self-Adjusting DOM](https://blog.janestreet.com/self-adjusting-dom-and-diffable-data/)
- [Seven Implementations of Incremental (Talk)](https://www.janestreet.com/tech-talks/seven-implementations-of-incremental/)
- [McCauley et al. — Incremental Topological Ordering with Predictions (ICML 2024)](https://proceedings.mlr.press/v235/mccauley24a.html)
- [Parallel Topological Sorting (CMU 2024)](https://github.com/codeplay0314/parallel-topological-sorting)
- [V8 Hidden Classes and Maps](https://v8.dev/docs/hidden-classes)
- [JS Engine Fundamentals: Shapes and Inline Caches](https://mathiasbynens.be/notes/shapes-ics)
- [Bun FFI Reference](https://bun.sh/reference/bun:ffi)
- [Bun SQLite](https://bun.sh/docs/runtime/sqlite)
- [MCP Best Practices](https://mcp-best-practice.github.io/mcp-best-practice/best-practice/)
- [Designing APIs for AI Agents (MCP)](https://dev.to/codimow/designing-apis-for-ai-agents-first-what-the-model-context-protocol-taught-us-2j46)
- [7 Principles for AI Agent Tool Design](https://dev.to/alexchen31337/7-principles-for-ai-agent-tool-design-from-claude-code-real-world-systems-3dcd)
- [Generational Arena Allocator (JS)](https://github.com/richardanaya/generational-arena)
- [HAMT for JavaScript](https://github.com/mattbierner/hamt)
- [Immutable Collections for TypeScript](https://immutable-collections.seedtactics.com/)
- [@thi.ng/sparse CSR](https://docs.thi.ng/umbrella/sparse/classes/CSR.html)
- [rbush R-tree](https://github.com/mourner/rbush)
- [StAX-XML Streaming Parser](https://clickin.github.io/stax-xml/)
- [foxglove/mcap Zero-Copy TS Optimization](https://github.com/foxglove/mcap/pull/1185)
- [TurboSheets Architecture](https://turbosheets.notion.site/)
- [Univer Formula Engine Architecture](https://docs.univer.ai/guides/recipes/architecture/formula)
- [Univer vs Luckysheet Comparison](https://blog.univer.ai/posts/univer-vs-luckysheet-a-comprehensive-comparison-of-open-source-spreadsheet-solutions/)
- [Throbol: Fast Recalculation](https://throbol.com/post/speed)
- [Puncalc: Task-Based Parallelism in Spreadsheets (J. Supercomputing 2019)](https://link.springer.com/article/10.1007/s11227-019-02823-8)
- [Excel Smart Recalculation Engine (Decision Models)](https://www.decisionmodels.com/calcsecrets.htm)
- [Excel Multithreading and Memory Contention](https://learn.microsoft.com/en-us/office/client-developer/excel/multithreading-and-memory-contention-in-excel)
- [Trueno: SIMD+GPU Compute Library](https://crates.io/crates/trueno)
- [Formualizer vs HyperFormula Comparison](https://docs.bswen.com/blog/2026-03-04-formualizer-vs-hyperformula-comparison/)
