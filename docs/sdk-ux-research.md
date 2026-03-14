# SDK & Agent UX Research: Actionable Insights for Ascend

Research on AI/agent-friendly spreadsheet interfaces, developer ergonomics, and modern TypeScript API design. Focus is on concrete, implementable improvements to Ascend's SDK, MCP server, and developer surface.

---

## 1. Agent/AI UX for Spreadsheets

### 1.1 The Core Problem: Token Budgets

A 576×23 spreadsheet naively serialized to JSON produces ~61,000 tokens. MCP tool schemas themselves can consume 20-40% of the agent's context window — a MySQL MCP server with 106 tools consumed 54,600 tokens just on initialization. Every token spent on tool definitions or raw data dumps is a token not spent on reasoning.

**Ascend's current position is strong.** The `agentView` / `AgentViewResult` already provides column-level summaries, formula pattern analysis, and sample rows. The `MachineEnvelope` gives structured `ok`/`error` responses with error codes, refs, and `suggestedFix`. These are ahead of most spreadsheet MCP servers that just dump raw cell data.

### 1.2 Ideal Agent Interface: Discovery → Profile → Extract

The best-performing pattern from production MCP spreadsheet servers (spreadsheet-kit, spreadsheet-read-mcp) follows a three-phase workflow:

| Phase | Purpose | Token Cost | Ascend Equivalent |
|-------|---------|------------|-------------------|
| **Discover** | Orient the agent: what sheets, tables, names exist | ~200-500 tokens | `ascend.inspect` (exists) |
| **Profile** | Understand structure without reading data: column types, cardinality, formula patterns | ~500-2000 tokens | `ascend.agent_view` (exists) |
| **Extract** | Surgical reads of specific ranges, filtered and paginated | Variable | `ascend.read` (exists, but see gaps below) |

**What Ascend is missing in this flow:**

1. **No `find_value` / search tool.** Agents frequently need to locate specific values, headers, or labels before they can construct range references. Currently an agent must read an entire sheet or guess ranges. A `ascend.find` tool that searches for values/patterns within a sheet (returning matching cell refs) would save significant back-and-forth.

2. **No table-aware read in MCP.** The SDK has `TableHandle` and `readTableWindow`, but no MCP tool exposes table-level reads. Agents working with structured tables shouldn't need to know the underlying A1 range — `ascend.read_table` with table name + row offset/limit would be more natural.

3. **No formula evaluation tool.** An agent can read existing formula results but can't ask "what would `=VLOOKUP(...)` return in this context?" without writing to a cell. A stateless `ascend.eval` tool that evaluates a formula expression against the current workbook state would enable exploratory analysis.

### 1.3 Context Window Optimizations

**What Ascend already does well:**
- `AgentViewResult` compresses a range into column summaries with kind, sample values, and formula patterns
- `CompactCellInfo` omits refs when unnecessary
- `flatValues` option strips `CellValue` wrappers to raw primitives
- Windowed reads with `rowOffset`/`rowLimit` and `hasMore`/`nextRowOffset` pagination

**What should change:**

| Optimization | Impact | Effort |
|-------------|--------|--------|
| **TSV-format output mode for `ascend.read`** | 3-5x token reduction vs JSON cell arrays. Tabular data as TSV is the most token-efficient representation LLMs can parse. A 100-row × 10-col read as JSON objects: ~8,000 tokens. Same data as TSV: ~1,500 tokens. | Low — add a `format: 'tsv'` option that returns `{ text: string, rowCount: number, hasMore: boolean }` |
| **Column-pruning in reads** | Agents often need 3-5 columns from a 20-column sheet. `cols` parameter accepting column letters or header names would avoid returning unwanted data. | Low |
| **Diff-aware reads with change tokens** | `CompactRangeWindowInfo` already has `changeToken`. The MCP tool should accept a `changedSince` token so repeated reads only return deltas. Critical for agent loops that repeatedly check workbook state. | Medium — `changedSince` already exists in `AgentReadOptions`, wire it through MCP |
| **Auto-summarize large results** | When a read would exceed ~4,000 tokens, automatically truncate and append a note: "Showing 50 of 2,340 rows. Use rowOffset/rowLimit to paginate." Currently the agent gets back a massive JSON blob and may not notice pagination fields. | Low |

### 1.4 MCP Tool Design Patterns

**Schema token efficiency is critical.** Every tool description and parameter schema is injected into the agent's context. Current Ascend MCP tool descriptions are already concise, which is good.

Best practices from the MCP ecosystem:

1. **Use enums, not free-text strings.** Ascend's `format: z.enum(['cells', 'rows', 'objects'])` is correct. Extend this to operation types.

2. **Return stable IDs, not full objects.** For large result sets, return row indices or cell refs that can be used in follow-up calls rather than embedding all data.

3. **Use `oneOf` for mutually exclusive options.** Instead of accepting both `range` and `table` parameters (where one would be ignored), use discriminated schemas.

4. **Resources for static metadata.** Sheet lists, defined names, and style summaries don't change between reads. Expose these as MCP Resources (read-only, cacheable) rather than tool calls. The MCP spec distinguishes Resources (knowledge) from Tools (actions).

5. **Server-level `instructions` field.** Add a top-level instruction to the MCP server that describes the recommended workflow (inspect → profile → read → write). This guides agent tool selection before individual tool schemas are evaluated.

**Concrete MCP tool schema improvement — `ascend.write` operations:**

The current `ops: z.array(z.record(z.string(), z.unknown()))` is opaque. Agents can't discover valid operation shapes from the schema. Replace with discriminated union schemas:

```typescript
const setCellOp = z.object({
  op: z.literal('set-cell'),
  sheet: z.string(),
  ref: z.string().describe('Cell reference like "A1"'),
  value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
  formula: z.string().optional().describe('Formula like "=SUM(A1:A10)"'),
})

const insertRowOp = z.object({
  op: z.literal('insert-row'),
  sheet: z.string(),
  row: z.number().int().positive(),
  count: z.number().int().positive().default(1),
})

const operation = z.discriminatedUnion('op', [setCellOp, insertRowOp, /* ... */])
```

This trades schema token cost for dramatically better agent accuracy. Per research, 68% of agent tool-use errors come from incorrect parameter values — typed schemas prevent this.

### 1.5 Error Messages for Agent Consumption

Ascend's `AscendError` is already well-structured with `code`, `message`, `retryable`, `refs`, `details`, and `suggestedFix`. This is better than 90% of MCP servers.

**Improvements for agent reliability:**

| Field | Current State | Recommended Change |
|-------|--------------|-------------------|
| `suggestedFix` | Optional, rarely populated | **Always populate for recoverable errors.** E.g., `SHEET_NOT_FOUND` → `"Available sheets: Sheet1, Data, Summary"`. `INVALID_REF` → `"Did you mean 'A1:C10'? Valid range format: SheetName!A1:C10"` |
| `retryable` | Boolean | Add `retryStrategy` — `'same'` (transient, retry as-is), `'modified'` (fix the input and retry), `'none'` (unrecoverable). Agents need to know whether to retry the same call or modify parameters. |
| Error context | `refs` array | For formula errors, include `precedents` — the cells that fed into the failed formula. Agents debugging formula chains need the upstream context. |
| Validation errors | Single error returned | Return **all** validation errors in a batch, not just the first. An agent sending 10 operations shouldn't need 10 round-trips to discover each error. |

### 1.6 Metadata: Automatic vs. On-Demand

| Auto-surface (always include) | On-demand (only when requested) |
|-------------------------------|--------------------------------|
| Sheet names + counts | Full sheet metadata (merges, conditional formats, protection) |
| Used range per sheet | Cell-level style information |
| Table names + column headers | Formula ASTs and token arrays |
| Defined name list | Dependency traces |
| Load mode + partial status | Workbook views, page setup |
| Compatibility warnings | Image/drawing references |
| Error summary (count by type) | Individual cell comments |

The `ascend.inspect` tool should return the "auto-surface" set by default without requiring `sheet` parameter. This gives agents enough context to decide what to drill into.

---

## 2. Developer UX for Spreadsheet SDKs

### 2.1 API Comparison: What Makes SDKs Ergonomic

| Library | Paradigm | Strength | Pain Point |
|---------|----------|----------|------------|
| **SheetJS** | Utility functions + plain objects | Zero-dependency, fast reads | `XLSX.utils.book_new()` is un-idiomatic; no method chaining; paid Pro for styling |
| **ExcelJS** | OOP class hierarchy (Workbook→Row→Cell) | Intuitive construction | OOM on large files; per-cell style objects; unreliable streaming |
| **OpenPyXL** | OOP with modal read/write-only | Explicit memory modes | 50x memory in standard mode; cell-by-cell access is 16x slower than generators |
| **Apache POI** | Enterprise OOP (XSSF/SXSSF/HSSF) | Most complete Excel compatibility | 630 bytes/cell; streaming breaks cross-window formulas |
| **Ascend SDK** | Read-view + mutable workbook | Agent-friendly summaries; preservation capsules; formula engine | See gaps below |

**What developers actually complain about (from GitHub issues, Stack Overflow, npm-compare):**

1. **"I just want to read this as JSON."** The most common use case — read XLSX, get array of objects with headers as keys. SheetJS's `sheet_to_json` handles this but requires understanding CSF internals. Ascend's `readObjects` with `headers: 'first-row'` is the right pattern.

2. **"Styling is impossible."** ExcelJS users report 20+ second style processing times. SheetJS gates styling behind a Pro license. Ascend's `StyleRegistry` avoids per-cell duplication, which is the correct architecture.

3. **"Large files crash my process."** Both SheetJS and ExcelJS fail on files >50MB without careful memory management. Ascend's selective sheet loading (`mode: 'metadata-only'`, `sheets: ['specific']`) is a significant advantage.

4. **"The API doesn't match how I think."** SheetJS's A1-key addressing (`sheet["A1"]`) forces developers to construct cell addresses manually. ExcelJS's `worksheet.getCell('A1')` is marginally better. Ascend's `SheetHandle` and `TableHandle` abstractions are the right direction.

### 2.2 Ascend SDK Ergonomic Gaps

**Current pattern — creating and writing a workbook:**
```typescript
const wb = Ascend.create()
wb.apply([
  { op: 'set-cell', sheet: 'Sheet1', ref: 'A1', value: { kind: 'string', value: 'Name' } },
  { op: 'set-cell', sheet: 'Sheet1', ref: 'B1', value: { kind: 'number', value: 42 } },
])
```

The `{ kind: 'string', value: 'Name' }` wrapping is correct internally but verbose for the developer surface. Compare with the ideal:

```typescript
const wb = Ascend.create()
wb.cell('Sheet1!A1').set('Name')
wb.cell('Sheet1!B1').set(42)
wb.cell('Sheet1!C1').formula('=A1&B1')
```

**Recommended additions (not replacements — these should coexist with the operations API):**

1. **Fluent cell access:** `wb.cell('Sheet1!A1')` → returns a `CellHandle` with `.set(value)`, `.formula(str)`, `.clear()`, `.style({...})`. Auto-coerces primitives to `CellValue`.

2. **Sheet-scoped shorthand:** `wb.sheet('Sheet1').cell('A1').set('Name')` — avoids repeating sheet name.

3. **Batch builder:**
   ```typescript
   wb.batch()
     .set('Sheet1!A1', 'Name')
     .set('Sheet1!B1', 42)
     .formula('Sheet1!C1', '=A1&B1')
     .commit()
   ```
   This collects operations and applies them atomically with a single `apply()` call.

4. **Auto-coercion for `set-cell` values:** Accept `string | number | boolean | null | Date` and internally convert to `CellValue`. The operations API should accept both raw primitives and wrapped `CellValue` objects.

### 2.3 Read-Only Views vs. Mutable Workbooks

Ascend's `WorkbookReadView` / `AscendWorkbook` split is the right architectural decision. OpenPyXL validates this pattern — their read-only and write-only modes reduce memory by 90%.

**Improvements:**

1. **Type-level enforcement.** `WorkbookReadView` should not have any method that could accidentally lead to mutation. Currently it exposes `getWorkbookModel()` which returns the raw `Workbook` — a mutation escape hatch. Consider `Readonly<Workbook>` or a separate `ReadonlyWorkbook` type.

2. **`WorkbookDocument.open()` vs `Ascend.open()` naming.** Having two entry points (`WorkbookDocument` for read-only, `Ascend` for mutable) works, but the distinction should be clearer. Consider `Ascend.read(path)` → `WorkbookReadView` and `Ascend.open(path)` → `AscendWorkbook`, both on the same namespace.

### 2.4 Batch Operations and Transaction Semantics

**Current:** `wb.apply(ops)` accepts an array of operations and returns `ApplyResult` with `errors` and `affectedCells`. There's also a `_batchMode` flag and `batch(fn)` method.

**What's missing:**

1. **Preview before commit.** `wb.preview(ops)` exists and returns diffs — this is excellent. But the pattern should be `preview → confirm → apply`, not `apply → hope it worked`. The MCP server should guide agents through this flow.

2. **Partial failure semantics.** If operation 5 of 10 fails, are operations 1-4 applied? The current behavior should be documented explicitly. Recommendation: all-or-nothing by default, with an `opts.partial: true` flag for best-effort application.

3. **Undo support.** `apply()` could return an `undo` function or inverse operation set. Critical for agent workflows where the agent might need to roll back.

### 2.5 Formula Evaluation Error Reporting

**Current:** `RecalcResult.errors` is `ReadonlyArray<{ ref: string; error: AscendError }>`. This tells you *what* failed but not *why* in a way that enables debugging.

**Ideal formula error report:**

```typescript
interface FormulaEvalError {
  ref: string                          // "Sheet1!C5"
  formula: string                      // "=VLOOKUP(A5,Data!A:C,3,FALSE)"
  errorType: string                    // "#N/A"
  message: string                      // "VLOOKUP: no match found for value 'XYZ'"
  evaluationTrace?: {                  // opt-in
    step: string                       // "VLOOKUP argument 1 → 'XYZ'"
    intermediateValues?: Record<string, unknown>
  }[]
  precedentValues?: Record<string, unknown>  // { "A5": "XYZ", "Data!A:C": "<range 100 rows>" }
  suggestedFix?: string                // "Check that 'XYZ' exists in column A of the Data sheet"
}
```

This gives agents (and developers) enough context to diagnose and fix formula errors without additional round-trips.

### 2.6 Streaming Large Workbooks

**Current:** The SDK has `streamWindowsCompact` (generator-based windowed iteration). The MCP server doesn't expose streaming — each `ascend.read` call opens and re-parses the file.

**Gaps:**

1. **Session-based MCP reads.** The `WorkbookDocument` / session cache already exists. The MCP server should use `configureSessionCache()` so repeated reads of the same file don't re-parse. This is a critical performance issue for agent workflows that make 5-10 read calls against the same file.

2. **Streaming write not exposed.** For agents building large workbooks, there's no way to stream rows in. A `ascend.append_rows` tool would enable building large datasets without holding the entire workbook in the agent's context.

3. **Progress feedback for large operations.** When recalculating a 100K-cell workbook, the agent should get progress indicators, not a 30-second timeout. MCP supports `notifications` — use them for long-running operations.

---

## 3. Modern TypeScript API Design

### 3.1 Patterns from Zod, Drizzle, tRPC, Effect

The unifying principle: **types are the API, not documentation.**

| Pattern | Example | Applicable to Ascend |
|---------|---------|---------------------|
| **Schema = type + validation** | Zod: `z.string().email()` → runtime check + inferred `string` type | Operation schemas should be Zod schemas that both validate at runtime and provide TypeScript types. Currently operations are `z.record(z.string(), z.unknown())` in MCP. |
| **Infer, don't declare** | Drizzle: table schema → inferred insert/select types | `CellValue` kind discriminant should drive type narrowing: `if (cell.kind === 'number') { cell.value /* number */ }` — this already works via discriminated unions. |
| **End-to-end type flow** | tRPC: server procedure → client gets exact return type | MCP `structuredContent` should carry type information that client-side tooling can leverage. |
| **Pipe/compose** | Effect: `Effect.gen(function*() { ... })` chains | Formula evaluation could expose a pipe-style API for composing transformations on ranges. |

### 3.2 Builder Pattern vs. Functional Composition vs. Method Chaining

For Ascend specifically:

| Pattern | Best For | Ascend Use Case |
|---------|---------|-----------------|
| **Method chaining** (return `this`) | Linear configuration, read-like queries | `wb.sheet('S1').range('A1:C10').readRows()` — natural for reads |
| **Builder** (accumulate, then build) | Complex object construction with validation | `OperationBuilder` for constructing multi-step edits, validating before commit |
| **Functional composition** (pipe transforms) | Data transformation pipelines | Range transforms: `range.filter(r => r.age > 30).map(r => r.name).toArray()` |

**Recommendation:** Method chaining for the primary read API, builder for write operations, functional composition for data transforms. Don't pick one — use each where it fits.

### 3.3 Type-Level Autocompletion

**Current gap:** `wb.cell('Sheet1!A1')` takes a plain `string`. TypeScript can't help with valid sheet names or cell refs.

**Practical improvements (no template literal type abuse):**

1. **Overloads for common patterns:**
   ```typescript
   cell(ref: string): CellHandle
   cell(sheet: string, row: number, col: number): CellHandle
   cell(sheet: string, a1Ref: string): CellHandle
   ```

2. **Narrowing on read results:**
   ```typescript
   const val = wb.cellValue('A1')
   if (val.kind === 'number') {
     val.value // TypeScript knows this is number
   }
   ```
   This already works with discriminated unions in `CellValue`.

3. **Typed operation builders:**
   ```typescript
   wb.ops.setCell({ sheet: 'Sheet1', ref: 'A1', value: 42 })
   //                  ^-- autocomplete sheet names from a generic?
   ```
   Sheet name autocomplete isn't practical at the type level (sheets are runtime data), but operation shape autocomplete via typed builders is.

### 3.4 Optional Configuration

Zod's `.optional().default()` pattern is the gold standard. For Ascend:

```typescript
// Current — fine but could be cleaner
Ascend.open(path, { mode: 'values', sheets: ['Sheet1'] })

// Consider named presets for common patterns
Ascend.open(path)                    // full mode (default)
Ascend.open(path, 'metadata-only')   // string shorthand for common modes
Ascend.open(path, { mode: 'selective', sheets: ['Sheet1'] })  // full options
```

Accept `string | OptionsObject` where the string maps to a preset configuration.

---

## 4. Documentation and Discoverability

### 4.1 Self-Documenting SDK

The strongest signal of a well-designed SDK is that developers can use it productively through autocompletion alone, without opening docs. This requires:

1. **Method names that describe the return value.** `inspect()` → workbook metadata, `readRows()` → row arrays, `readObjects()` → keyed objects. Ascend's naming is already good here.

2. **JSDoc on every public method.** Include the return type shape, not just a description. Example:
   ```typescript
   /**
    * Read a range as arrays of cell values, one array per row.
    * Returns `undefined` if the sheet doesn't exist.
    *
    * @example
    * const result = wb.sheet('Sales').readRows('A1:D100', { rowLimit: 50 })
    * result.rows[0] // [{ kind: 'string', value: 'Name' }, ...]
    */
   readRows(range: string, opts?: AgentReadOptions): RangeRowsInfo | undefined
   ```

3. **Inline examples in JSDoc.** The `@example` tag appears in VS Code hover popups. This is the most discoverable documentation surface — more visible than README or docs site.

4. **Return type naming convention.** Ascend uses `*Info` suffix (`SheetInfo`, `TableInfo`, `RangeWindowInfo`). This is consistent and good. Keep it.

### 4.2 JSDoc vs. Runtime Validation

Don't choose — use both. JSDoc for development-time documentation and type checking. Runtime validation (via Zod or manual checks) for SDK boundaries where untrusted input enters.

**Where to validate at runtime:**
- `Ascend.open(path, options)` — validate options shape
- `wb.apply(ops)` — validate operation array
- MCP tool inputs — already validated via Zod schemas
- Any public method accepting user-provided refs or ranges — validate A1 format

**Where JSDoc alone is sufficient:**
- Internal methods called by the SDK itself
- Return types (if the SDK constructs them, they're valid by construction)
- Generic type parameters

### 4.3 Example Structure

From the most successful TypeScript SDKs (Stripe, Prisma, Drizzle), examples should follow this pattern:

1. **Copy-pasteable.** Every example should work if pasted into a file with the import added. No pseudocode.

2. **Show the result.** Include a comment showing what the code returns:
   ```typescript
   const info = wb.inspect()
   // { sheetCount: 3, sheets: [{ name: 'Sheet1', rowCount: 100, ... }], ... }
   ```

3. **Progressive complexity.** Start with the 1-line version, then show options:
   ```typescript
   // Simple
   const wb = await Ascend.open('data.xlsx')

   // With options
   const wb = await Ascend.open('data.xlsx', {
     mode: 'values',
     sheets: ['Summary'],
   })
   ```

4. **Error handling shown explicitly.** Don't only show the happy path:
   ```typescript
   const result = wb.apply(ops)
   if (result.errors.length > 0) {
     for (const err of result.errors) {
       console.error(`${err.code}: ${err.message}`)
       if (err.suggestedFix) console.log(`  Fix: ${err.suggestedFix}`)
     }
   }
   ```

---

## 5. Priority Action Items for Ascend

Ranked by impact × feasibility:

### Immediate (< 1 week each)

1. **Add TSV output mode to `ascend.read`.** Accept `format: 'tsv'` that returns tab-separated text instead of JSON cell arrays. 3-5x token savings for agents. Trivial to implement — the CSV writer already handles TSV.

2. **Add `ascend.find` MCP tool.** Search for values/patterns within a sheet. Returns matching cell refs. Prevents agents from reading entire sheets to locate data.

3. **Populate `suggestedFix` on all recoverable errors.** Every `SHEET_NOT_FOUND` should list available sheets. Every `INVALID_REF` should show the expected format. Every `FORMULA_PARSE_ERROR` should indicate what token was unexpected.

4. **Wire `changedSince` through `ascend.read` MCP tool.** The SDK already supports change tokens. Expose them so agent loops don't re-read unchanged data.

5. **Type the operations schema in MCP.** Replace `z.record(z.string(), z.unknown())` with discriminated union Zod schemas for each operation type. Biggest single improvement for agent write reliability.

### Short-term (1-2 weeks each)

6. **Add `ascend.read_table` MCP tool.** Read table data by table name with row pagination. Tables are the natural unit for structured data — agents shouldn't need to know underlying ranges.

7. **Fluent cell/range API on `AscendWorkbook`.** `wb.cell('Sheet1!A1').set(42)` as sugar over the operations API. Auto-coerce primitives.

8. **Batch builder API.** `wb.batch().set(...).set(...).commit()` — collect operations, apply atomically.

9. **Session caching in MCP server.** Use `configureSessionCache()` so repeated operations on the same file don't re-parse. Order-of-magnitude speedup for agent workflows.

10. **MCP Resources for static metadata.** Expose sheet list, defined names, and style summary as MCP Resources (cacheable, no tool call needed).

### Medium-term (2-4 weeks)

11. **Richer formula error diagnostics.** Include formula text, error type (#N/A, #REF, etc.), precedent values, and evaluation trace in `RecalcResult.errors`.

12. **`ascend.eval` tool.** Stateless formula evaluation — "what would this formula return?" without mutating the workbook. Enables agent exploration.

13. **Undo/inverse operations from `apply()`.** Return the inverse operation set so agents can roll back.

14. **Column-pruning in reads.** Accept `cols: ['A', 'C', 'F']` or `cols: ['Name', 'Amount']` (header-based) to limit returned columns.

15. **JSDoc + examples on every public SDK method.** With `@example` tags that show up in editor hover.

---

## References

- [Token-Efficient Agents: MCP-Heavy Agents Without Burning Tokens](https://codeagentsalpha.substack.com)
- [MCP Tool Schema Bloat: The Hidden Token Tax](http://layered.dev/mcp-tool-schema-bloat)
- [Writing Great Tool Schemas for MCP](https://www.mcpbundles.com/blog/2025/05/06/writing-great-tool-schemas)
- [spreadsheet-read-mcp — Token-efficient XLSX reads](https://crates.io/crates/spreadsheet-read-mcp)
- [SheetAgent: Generalist Agent for Spreadsheet Manipulation](https://sheetagent.github.io/)
- [Beyond Rows to Reasoning: Agentic Retrieval for Spreadsheets (arxiv 2603.06503)](https://arxiv.org/html/2603.06503v1)
- [How to Fit Massive Excel Files into LLMs: Compression Playbook](https://medium.com/@denisuraev)
- [Advanced Function Calling & Tool Composition for Production Agents (2026)](https://iterathon.tech/blog/advanced-function-calling-tool-composition-production-agents-2026)
- [MCP Best Practice](https://mcp-best-practice.github.io/mcp-best-practice/best-practice/)
- [MCP Features Guide — WorkOS](https://workos.com/blog/mcp-features-guide)
- [The New Wave of TypeScript-First Libraries in 2026 — PkgPulse](https://www.pkgpulse.com/blog/new-wave-typescript-first-libraries-2026)
- [FxD: A Functional Debugger for Spreadsheets](https://www.iandrosos.me/fxd.html)
- [ExceLint: Automatically Finding Spreadsheet Formula Errors (Microsoft Research)](https://www.microsoft.com/en-us/research/publication/excelint-automatically-finding-spreadsheet-formula-errors/)
- [SheetJS vs ExcelJS API Comparison](https://npm-compare.com/exceljs,xlsx)
- [SheetJS Idiomatic API Discussion (Issue #1331)](https://github.com/SheetJS/sheetjs/issues/1331)
