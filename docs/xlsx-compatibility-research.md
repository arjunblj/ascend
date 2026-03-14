# XLSX Compatibility & Test Fixtures Research

What real-world XLSX files contain, what patterns matter most, where to find test fixtures, and every known gotcha for implementing Excel compatibility.

---

## 1. Test Fixture Sources

### Tier 1: Essential (use immediately)

| Source | URL | Coverage | License/Access | Comprehensiveness |
|--------|-----|----------|----------------|-------------------|
| **SheetJS test_files** | `https://oss.sheetjs.com/test_files/` (ZIP: `https://test-files.sheetjs.com/test_files.zip`) | AutoFilter, blank sheet types, number format conditions, RK numbers, comments stress test, custom properties, formula stress test, LONumbers. 680+ files across XLS/XLSX/XLSB/XML. Organized by Excel version (1904, 2011, 2013, 2016). | Apache 2.0. GitHub repo currently disabled (TOS violation) but files hosted directly. | **Very high.** Best single-source collection for format-level edge cases. |
| **Apache POI test-data/spreadsheet** | `https://github.com/apache/poi/tree/trunk/test-data/spreadsheet` | Protection (workbook/sheet/revision), shared formulas, SUMIF/SUMIFS, array formulas, conditional formatting, cell styles, Unicode names, rich text, autofilter, VML drawing, named ranges, XOR encryption, charts, YEARFRAC, merged cells. Hundreds of files. | Apache 2.0. | **Very high.** Gold standard for edge cases — each file typically represents a specific bug or feature test. File names reference issue numbers for traceability. |
| **XlsxWriter comparison tests** | `https://github.com/jmcnamara/XlsxWriter/tree/master/xlsxwriter/test/comparison` | Array formulas, autofilters, headers/footers with images, merged cells, charts, data validation, conditional formatting, page setup, print areas. Each test has a reference XLSX created by Excel + a Python script that generates the equivalent. | BSD. | **High.** Unique value: pairs of (Excel-generated reference, programmatic reproduction) make them ideal for binary comparison testing. |

### Tier 2: Valuable (incorporate selectively)

| Source | URL | Coverage | License/Access | Comprehensiveness |
|--------|-----|----------|----------------|-------------------|
| **Calamine tests** | `https://github.com/tafia/calamine/tree/master/tests` | Read-path validation: cell types, formulas (preserved strings), dates, number formats, merged cells, defined names. Focus on read correctness across XLS/XLSX/XLSB/ODS. | MIT. | **Medium.** Read-only focus, but good for validating parser correctness. |
| **OpenPyXL tests** | `https://foss.heptapod.net/openpyxl/openpyxl` (also `https://github.com/openpyxl/openpyxl-ci`) | Rich text, cell styles, conditional formatting, charts, comments, data validation, page setup, named ranges, print areas. Tests in `openpyxl/tests/`. | MIT. | **Medium-high.** Strong on style/formatting edge cases. Specific date handling test repo at `github.com/eawag-rdm/openpyxl_date-handling-test`. |
| **HyperFormula** | `https://github.com/handsontable/hyperformula` | 350/515 Excel functions (68% coverage), 5,000+ unit tests at 97% code coverage. Configuration-based Excel compatibility (string comparison, wildcard criteria, date/time formats). | GPLv3 (tests usable as reference, not embeddable). | **Very high** for formula evaluation correctness. Best open-source formula conformance suite. |
| **ExcelNumberFormat** | `https://github.com/andersnm/ExcelNumberFormat` | Comprehensive Excel number format string parser. Handles all 4 sections (positive/negative/zero/text), locale codes, date format tokens, conditional format thresholds, color codes, fill characters. | MIT. C# but algorithms are portable. | **High** for number format parsing specifically. |
| **XLParser** | `https://github.com/spreadsheetlab/XLParser` | Formula parsing with 99.9% compatibility rate against real-world formulas. Corpus-tested. | MIT. C#. | **High** for parser edge cases and grammar coverage. |
| **libxlsxwriter** | `https://github.com/jmcnamara/libxlsxwriter` (same author as XlsxWriter) | 749 functional tests + 429 unit tests. C library. Same comparison methodology as XlsxWriter. | BSD. | **High.** The C implementation often catches edge cases the Python version doesn't. |

### Tier 3: Reference (consult for specific questions)

| Source | URL | Coverage | Notes |
|--------|-----|----------|-------|
| **OOXML Validator** | `https://github.com/mikeebowen/OOXML-Validator` | .NET CLI that validates XLSX against Office 2007–365 schemas. | Useful for validating our write output. |
| **ISO/IEC 29500 Schemas** | `https://github.com/sc34wg4/OOXMLSchemas` | Official XSD schemas for SpreadsheetML. | Normative reference. |
| **MS-OI29500** | `https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500` | Microsoft's documentation of where Office diverges from ISO 29500. | Critical for understanding real-world deviations from spec. |
| **office-interoperability-tools** | `https://github.com/x1sc0/office-interoperability-tools` | Cross-application interop testing. | Niche but useful for multi-app compat. |
| **POI regression corpus** | `https://github.com/centic9/poi-regression-test` | Runs POI against a massive corpus of real-world files. | Methodology reference for large-scale regression testing. |
| **tidyxl (R)** | `https://rdrr.io/cran/tidyxl/src/tests/testthat/test-compatibility.R` | Tests reading XLSX from Excel, GnuMeric, and other generators. | Good cross-generator compat patterns. |

---

## 2. What Real-World Excel Files Actually Contain

### Feature Usage Frequency (from surveys and corpus analysis)

**Tier 1 — Used in >80% of real workbooks:**
- Basic cell values: numbers, strings, dates, booleans
- Number formatting (currency, percentage, date formats, decimal places)
- Cell styles: fonts, fills, borders, alignment
- SUM, AVERAGE, COUNT, IF, VLOOKUP/XLOOKUP
- Merged cells
- Multiple sheets
- Column widths, row heights
- Freeze panes

**Tier 2 — Used in 30-80% of workbooks:**
- Conditional formatting (color scales, data bars, icon sets, formula-based rules)
- Named ranges (workbook-scoped)
- Data validation (dropdowns, input constraints)
- SUMIF/SUMIFS, COUNTIF/COUNTIFS, INDEX/MATCH
- Charts (embedded)
- AutoFilter
- Print area / page setup
- Comments / notes
- Hyperlinks

**Tier 3 — Used in 5-30% of workbooks:**
- Array formulas (CSE and dynamic arrays)
- IFERROR, INDIRECT, OFFSET
- Pivot tables (read-only is fine)
- Protection (sheet/workbook)
- Rich text in cells (mixed formatting within a cell)
- Custom number formats (complex multi-section)
- Named ranges (sheet-scoped)
- External references
- Tables (structured references)

**Tier 4 — Used in <5% but critical for power users:**
- LAMBDA, LET, XLOOKUP, FILTER, SORT, UNIQUE, SEQUENCE
- Dynamic array spill behavior
- Implicit intersection (@)
- R1C1 notation
- Defined names with formulas (not just ranges)
- VBA/macros (preserve, don't execute)
- Sparklines
- Slicers

### The "12 Functions That Drive 97% of Financial Models"

Research analyzing 100+ professional financial models found that just 12 functions account for ~97% of all formulas:

1. SUM
2. IF
3. VLOOKUP / INDEX+MATCH
4. SUMIF / SUMIFS
5. COUNT / COUNTA
6. AVERAGE
7. ROUND
8. MAX / MIN
9. LEFT / RIGHT / MID
10. CONCATENATE / TEXTJOIN
11. IFERROR
12. DATE / YEAR / MONTH / DAY

---

## 3. Known Gotchas and Edge Cases

### 3.1 Date Serial Numbers — The 1900 Leap Year Bug

**The bug:** Excel treats 1900 as a leap year, allowing serial number 60 = Feb 29, 1900 (which never existed). This was inherited from Lotus 1-2-3 for backward compatibility.

**Impact on implementation:**
- Serial numbers 1–59 map to Jan 1–Feb 28, 1900
- Serial number 60 = the fictional Feb 29, 1900
- Serial numbers 61+ are all off by 1 day from the true Gregorian calendar
- `WEEKDAY()` returns wrong values for dates before March 1, 1900
- **All modern libraries must reproduce this bug** to maintain round-trip fidelity

**The 1904 date system:**
- macOS Excel historically used Jan 1, 1904 as epoch (avoids the 1900 bug)
- Serial numbers differ by 1,462 days between systems
- Workbooks store which system they use in `workbook.xml` (`<workbookPr date1904="1"/>`)
- Copying between workbooks with different date systems requires automatic offset
- Ascend already handles this via `dateSystem: '1900' | '1904'` in `FunctionEvalContext`

**Test cases needed:**
- Serial 0 = Jan 0, 1900 (Excel displays as "0" or "1/0/1900")
- Serial 1 = Jan 1, 1900
- Serial 59 = Feb 28, 1900
- Serial 60 = Feb 29, 1900 (fictional — must reproduce)
- Serial 61 = Mar 1, 1900
- Serial 44927 = Dec 31, 2022 (verify modern date serial)
- Negative serial numbers (Excel shows #NUM! or treats as text)
- Cross-system (1900↔1904) copy semantics

### 3.2 Number Format String Parsing

**Structure:** Up to 4 semicolon-separated sections: `positive;negative;zero;text`

**Key complexity:**
- 1 section → applies to all numbers
- 2 sections → positive+zero; negative
- 3 sections → positive; negative; zero
- 4 sections → positive; negative; zero; text
- Conditional thresholds: `[>1000]#,##0;[>100]#,##0.0;#,##0.00`
- Color codes: `[Red]`, `[Color3]`, `[DBNum1]`
- Locale codes: `[$-409]` (US English), `[$€-407]` (German Euro)
- Date/time tokens: `yyyy`, `mm`, `dd`, `hh`, `mm`, `ss`, `AM/PM` — note `mm` is ambiguous (month vs. minute, resolved by position relative to `hh`/`ss`)
- Fill character: `*` followed by any character to fill remaining width
- Spacing character: `_` followed by any character for invisible padding
- Literal text in quotes: `"Total: "#,##0`
- Escaped single chars: `\$`, `\-`
- Fraction formats: `# ?/?`, `# ??/??`, `# ?/8`
- Scientific: `0.00E+00`
- Max 255 characters

**Built-in format IDs that differ from ECMA-376:**
- IDs 0–49 are standardized but Excel deviates from spec for several (documented in MS-OI29500)
- Locale-dependent built-in formats (date formats differ by locale even for the same ID)

**Test cases needed:**
- Every built-in format ID (0–49) with sample values
- Multi-section formats with all 4 sections
- Conditional thresholds (`[>100]`, `[<=0]`)
- Date format ambiguity (`mm` as month vs. minute)
- Fraction formatting (denominators, alignment)
- Color codes (named and numeric)
- Locale-specific format IDs
- Fill and spacing characters
- Escaped literals and quoted text

### 3.3 Shared Formulas

**What they are:** When Excel saves, it optimizes storage by writing a formula once for a "master" cell and marking dependent cells as shared, storing only the reference offset.

**XLSX representation:**
```xml
<c r="B2"><f t="shared" ref="B2:B100" si="0">A2*2</f><v>4</v></c>
<c r="B3"><f t="shared" si="0"/><v>6</v></c>
<!-- B3's formula is A3*2, derived by adjusting B2's formula -->
```

**Gotchas:**
- The `ref` attribute on the master cell defines the range of the shared formula group
- Slave cells only have `si` (shared index), no formula text — you must reconstruct by adjusting relative references
- If any cell in the shared range has been individually edited, it may have its own formula text that overrides the shared formula
- Some generators write `ref` ranges that don't match the actual populated cells
- Shared formulas can span non-contiguous regions in malformed files
- R1C1 conversion is essential for correct shared formula expansion

**Test cases needed:**
- Basic shared formula expansion (column-wise, row-wise)
- Shared formula with mixed absolute/relative references
- Overridden cells within a shared formula range
- Shared formula across multiple sheets (should not happen, but malformed files exist)
- Empty cells within the shared formula ref range

### 3.4 Rich Text in Cells

**XLSX structure:** Rich text cells use `<si>` (shared string item) with multiple `<r>` (run) elements:
```xml
<si>
  <r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/></rPr><t>Bold Red</t></r>
  <r><rPr><i/><sz val="11"/></rPr><t> Italic Normal</t></r>
</si>
```

**Gotchas:**
- The first run may or may not have `<rPr>` — if absent, it inherits the cell's style
- Whitespace preservation: `<t xml:space="preserve"> text </t>` — without this attribute, leading/trailing spaces are stripped
- Rich text cells are stored in the shared string table, not inline
- Some generators put plain strings in `<si><t>simple</t></si>` (no runs), others always use `<r>` even for uniform formatting
- Font properties in `<rPr>` can include: bold, italic, underline, strikethrough, font name, font size, color (theme/rgb/indexed), vertical alignment (superscript/subscript), charset, font family
- Concatenating rich text cells in formulas strips formatting (formula results are always plain)

**Test cases needed:**
- Single run with formatting
- Multiple runs with different formatting
- Run without `<rPr>` (inherits cell style)
- Whitespace preservation with `xml:space`
- Rich text in shared string table vs. inline string
- Extremely long rich text cells (100+ runs)
- Rich text with theme colors and indexed colors

### 3.5 Conditional Formatting Evaluation

**Evaluation model:**
- Rules evaluated top-to-bottom as ordered in the Rules Manager
- Each format property (fill, font, border) set by the first matching rule "wins" — later rules can only set properties not yet claimed
- `stopIfTrue` attribute halts evaluation for a cell when a rule matches
- Rules can reference formulas that depend on other cells
- Rules can span non-contiguous ranges

**Types to handle:**
1. Cell value rules (`cellIs`): `greaterThan`, `lessThan`, `between`, `equal`, etc.
2. Formula rules: arbitrary formula that returns TRUE/FALSE
3. Color scales (2-color, 3-color): requires computing min/max/percentile of the range
4. Data bars: percentage fill relative to range min/max
5. Icon sets: threshold-based icon selection
6. Top/bottom N rules: requires sorting/ranking the range
7. Above/below average: requires computing range average
8. Duplicate/unique values: requires scanning the range
9. Text contains/begins with/ends with
10. Date-based rules (today, yesterday, this week, last month, etc.)
11. Blanks/no blanks, errors/no errors

**Gotchas:**
- Priority attribute in XLSX is 1-based and determines evaluation order
- A single `<conditionalFormatting>` element can contain multiple rules for the same range
- Multiple `<conditionalFormatting>` elements can overlap on the same cells
- Formula-based rules use relative references anchored to the top-left cell of the range
- Color scale/data bar computations must handle mixed types (ignore non-numeric)
- Conditional formats override manual formatting but manual formatting persists if rule is deleted

### 3.6 Implicit Intersection and Dynamic Arrays

**The @ operator:**
- In pre-dynamic-array Excel, formulas that returned arrays were silently intersected with the formula cell's row/column to produce a single value
- Dynamic array Excel makes this visible with the `@` prefix: `=@INDEX(A:A,1)` vs `=INDEX(A:A,1)` (which now spills)
- Files saved in older Excel versions get `@` inserted when opened in modern Excel — this is a display change, not a semantic one

**Spill behavior:**
- Dynamic array formulas (SORT, FILTER, UNIQUE, SEQUENCE, RANDARRAY) produce arrays that "spill" into adjacent cells
- Spill range is a single formula; only the anchor cell is editable
- `#SPILL!` error if spill range is obstructed
- `ANCHORARRAY` implicit reference: other formulas can reference a spill range by referring to the anchor cell with `#` suffix: `=SUM(A1#)`
- Spill ranges are not stored in the XLSX — they're computed on recalculation

**For Ascend's purposes:**
- Parse `@` operator in formulas (preserve in AST)
- Support spill for dynamic array functions (SORT, FILTER, UNIQUE, SEQUENCE, etc.)
- Implement `#SPILL!` error detection
- Support `#` anchor reference syntax

### 3.7 Volatile Functions

**Truly volatile (recalc on every change):**
- `NOW()`, `TODAY()`, `RAND()`, `RANDBETWEEN()`
- `OFFSET()`, `INDIRECT()`, `CELL()`, `INFO()`

**Observation-volatile (recalc when workbook opens):**
- These should be marked dirty when a workbook opens but don't need recalc on every cell edit

**Implementation concerns:**
- Volatile functions must be tracked in the dependency graph as always-dirty
- A single `NOW()` with 10,000 dependents causes 10,000 recalculations per edit
- Ascend marks volatile functions via `volatile?: boolean` in `FunctionDef` — verify this propagates correctly through the dep graph
- `INDIRECT()` and `OFFSET()` are volatile because their reference targets are dynamic — dep graph cannot statically resolve them
- Non-volatile alternatives exist (INDEX instead of OFFSET, direct references instead of INDIRECT) but real-world files use the volatile versions heavily

### 3.8 Unicode and Encoding

**XLSX internals are always UTF-8 XML.** No encoding issues within the ZIP. But:
- Sheet names can contain Unicode (Japanese, Arabic, emoji)
- Cell values in shared strings are UTF-8
- Formula text can reference sheets with Unicode names: `='日本語シート'!A1`
- Named ranges can have Unicode names
- File paths in external references may have Unicode
- Sorting/comparison should be locale-aware (Ascend uses `locale` in `FunctionEvalContext`)
- Some generators produce malformed UTF-8 (BOM in XML parts, incomplete surrogate pairs)

### 3.9 R1C1 vs. A1 Notation

**Storage:** XLSX always stores formulas in A1 notation. R1C1 is a display/authoring preference stored in `workbook.xml`.

**Conversion edge cases:**
- Column letters beyond Z: AA=27, AZ=52, AAA=703... (bijective base-26, not standard base-26)
- `$A$1` (absolute) = `R1C1` (no brackets)
- `A1` (relative) = `R[0]C[0]` (brackets = relative offset from formula cell)
- `$A1` (mixed) = `R[0]C1`
- `A$1` (mixed) = `R1C[0]`
- R1C1 relative references are essential for shared formula expansion

### 3.10 Shared String Table

**Architecture:**
- All unique strings stored in `xl/sharedStrings.xml`
- Cells reference strings by 0-based index
- Table must be fully loaded before any cell values can be resolved
- Rich text entries coexist with plain text entries in the same table

**Performance gotchas:**
- Large workbooks can have 1M+ unique strings → 100MB+ shared string XML
- Streaming read is impossible without random-access or pre-loading the table
- Index-based referencing means insertion order matters — the table is append-only during write
- Some generators emit duplicate entries (violating the "unique" contract) — readers must handle this
- `uniqueCount` attribute may not match actual count in malformed files
- Inline strings (`<is>` element in cell) bypass the shared table — both paths must work

### 3.11 Defined Names

**Types of defined names:**
1. **Range references:** `=Sheet1!$A$1:$D$100` (most common)
2. **Constants:** `=3.14159` or `="Header Text"`
3. **Formulas:** `=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)` (dynamic range)
4. **Hidden names:** `_xlnm.Print_Area`, `_xlnm.Print_Titles`, `_xlnm._FilterDatabase`
5. **External references:** `=[Budget.xlsx]Sheet1!$A$1`

**Scope rules:**
- Workbook scope (default): name is global
- Worksheet scope: name is local, can shadow workbook-level names
- Same name can exist at both levels — worksheet scope wins when referenced from that sheet
- Names can reference other names (creates dependency chains)
- Names can be self-referential (intentional for certain patterns, or errors)

**Hidden/reserved names:**
| Name | Purpose |
|------|---------|
| `_xlnm.Print_Area` | Print range for a sheet |
| `_xlnm.Print_Titles` | Repeated rows/columns for printing |
| `_xlnm._FilterDatabase` | AutoFilter range |
| `_xlnm.Criteria` | Advanced filter criteria |
| `_xlnm.Extract` | Advanced filter output |
| `_xlnm.Database` | Database range |
| `_xlnm.Sheet_Title` | Sheet title override |

### 3.12 The `_xlfn.` Prefix Problem

**What it is:** Functions introduced after Excel 2010 are prefixed with `_xlfn.` in the XML so older Excel versions show `#NAME?` rather than silently miscomputing:

```xml
<f>_xlfn.XLOOKUP(A1,B:B,C:C)</f>
<f>_xlfn.IFS(A1>90,"A",A1>80,"B")</f>
<f>_xlfn._xlws.FILTER(A1:A10,B1:B10>5)</f>
```

**The `_xlws.` sub-prefix:** Dynamic array worksheet functions get double-prefixed: `_xlfn._xlws.SORT`, `_xlfn._xlws.FILTER`, `_xlfn._xlws.UNIQUE`, etc.

**Implementation requirements:**
- Strip `_xlfn.` and `_xlfn._xlws.` when parsing formulas
- Re-add when writing XLSX (for backward compat with older Excel)
- Function lookup must work both with and without prefix
- Display to users should never show the prefix

### 3.13 Array Formulas (CSE vs. Dynamic)

**CSE (Ctrl+Shift+Enter) — legacy:**
```xml
<c r="A1"><f t="array" ref="A1:C3">{=A1:C3*2}</f></c>
```
- Fixed output range declared in `ref` attribute
- All cells in the range share the same formula
- Cannot be partially edited — the entire array must be entered/deleted as a unit

**Dynamic arrays (Excel 365+):**
- Single formula in anchor cell, spills automatically
- No `ref` attribute needed (or `ref` equals just the anchor cell)
- `cm="1"` attribute on the cell indicates dynamic array context
- Spill range is computed, not stored

**Gotchas:**
- CSE formulas store the result as a single value in each cell's `<v>` element
- Dynamic array formulas may or may not have cached results in the XML
- Some files have hybrid: CSE formula syntax with dynamic array metadata
- Empty cells in a CSE array range still belong to the array group

---

## 4. Prioritized Action Items

### A. Test Fixtures to Acquire/Create

**Priority 1 — Acquire immediately:**
1. Download SheetJS test_files ZIP (`https://test-files.sheetjs.com/test_files.zip`)
2. Clone Apache POI `test-data/spreadsheet/` directory
3. Clone XlsxWriter `xlsxwriter/test/comparison/xlsx_files/`

**Priority 2 — Create custom fixtures:**
4. **Date edge cases workbook:** Serial 0, 1, 59, 60, 61; 1904-system dates; negative serials; dates near 2100 (non-leap century year)
5. **Number format stress test:** All 50 built-in format IDs; 4-section custom formats; conditional formats; locale-specific formats; fraction formats; fill/spacing characters; ambiguous `mm`
6. **Shared formula workbook:** Column-wise shared, row-wise shared, mixed absolute/relative, overridden cells, gaps in shared range
7. **Rich text workbook:** Single run, multi-run, no-rPr first run, whitespace preservation, theme colors, 100+ runs per cell
8. **Conditional formatting workbook:** All 11 rule types; overlapping ranges; stopIfTrue; formula-based with relative refs; color scales with mixed types
9. **Defined names workbook:** Workbook vs sheet scope; name shadowing; formula-based names; hidden reserved names; self-referential names
10. **Dynamic arrays workbook:** SORT/FILTER/UNIQUE/SEQUENCE output; #SPILL! error; anchor references with #; mixed CSE and dynamic
11. **Unicode stress test:** Japanese/Arabic/emoji sheet names; Unicode in named ranges; Unicode in formulas; surrogate pairs; BOM edge cases
12. **Large shared string table:** 100K+ unique strings; duplicate entries; inline strings mixed with shared; rich text in shared table

**Priority 3 — Reference from existing suites:**
13. Mirror HyperFormula's 5,000+ formula test cases as JSON conformance fixtures
14. Port ExcelNumberFormat's test vectors to TypeScript
15. Use OOXML Validator to validate our write output against Office 2007–365 schemas

### B. Features for Real-World Compatibility (by priority)

**Must handle correctly (blocks adoption):**
1. Cell values: all types (number, string, boolean, date, error, empty, rich text)
2. Number formatting: built-in IDs + custom format string parsing/rendering
3. Formulas: the core 12 functions + IFERROR, INDIRECT, OFFSET, TEXT, CONCATENATE
4. Shared string table: read/write, handle duplicates, inline strings
5. Date serial numbers: 1900 bug, 1904 system, serial↔Date conversion
6. Cell styles: font, fill, border, alignment, number format (via StyleRegistry)
7. Merged cells: read/write, skip slave cells in recalc
8. Shared formulas: expansion with correct reference adjustment
9. Named ranges: workbook and sheet scope, resolution in formulas
10. Multiple sheets: cross-sheet references in formulas

**Should handle correctly (quality of life):**
11. Conditional formatting: at least cellIs, formula, and colorScale rules
12. Data validation: dropdowns, input constraints
13. AutoFilter: preserve range, filter state
14. Comments/notes: preserve text and author
15. Freeze panes: preserve split position
16. Column widths / row heights: preserve and apply
17. Print area / page setup: preserve
18. Hyperlinks: internal and external
19. `_xlfn.` prefix: strip on read, restore on write
20. Dynamic array functions: SORT, FILTER, UNIQUE, SEQUENCE (with spill)

**Nice to have (power user / full compat):**
21. Charts: preserve on read/write (no editing needed)
22. Pivot tables: preserve on read/write
23. Tables (structured references): ListObject, `[@Column]` syntax
24. Protection: sheet/workbook lock flags
25. VBA/macros: preserve in XLSM round-trip
26. Sparklines: preserve
27. R1C1 display preference: preserve workbook setting
28. Implicit intersection (@): parse and evaluate correctly

### C. Edge Cases That Trip Up Implementations

**Ranked by frequency of occurrence in bug trackers of POI, SheetJS, OpenPyXL, XlsxWriter:**

1. **Date serial 60 (Feb 29, 1900)** — Must reproduce the bug, not fix it
2. **`mm` ambiguity in number formats** — Month when near `d`/`y`, minute when near `h`/`s`
3. **Shared formula expansion with mixed references** — `$A1` vs `A$1` vs `$A$1` adjust differently
4. **Shared string table index off-by-one** — 0-based in XML, sometimes written as 1-based by broken generators
5. **`_xlfn.` prefix not stripped** — Causes #NAME? errors when evaluating formulas from XLSX
6. **Rich text whitespace loss** — Missing `xml:space="preserve"` drops leading/trailing spaces
7. **Conditional formatting priority order** — 1-based, higher number = lower priority (counterintuitive)
8. **Named range scope resolution** — Sheet-level name shadows workbook-level name only from that sheet
9. **Volatile function chain explosion** — Single `NOW()` with deep dependency chain causes full recalc
10. **Column letter conversion overflow** — Column 16384 (XFD) is the max; bijective base-26 is not standard base-26
11. **Empty string vs. empty cell** — `""` is type "string", empty cell is type "empty"; they compare differently in COUNTBLANK, IF, etc.
12. **Negative zero (-0)** — Excel treats -0 as 0 for display but preserves the sign bit; `1/(-0)` = `-Infinity`
13. **IEEE 754 precision** — `0.1 + 0.2 ≠ 0.3` in floating point; Excel rounds to 15 significant digits for display
14. **Array formula cell deletion** — Cannot delete a subset of a CSE array range
15. **Circular reference handling** — Excel supports intentional circulars with iteration settings; most engines just error
16. **Theme color resolution** — Color indices reference the workbook's theme, which can be customized; indexed colors (0–63) have a legacy palette
17. **Row/column limits** — XLSX supports 1,048,576 rows × 16,384 columns; XLS supports 65,536 × 256
18. **Calculation chain** — `xl/calcChain.xml` specifies recalculation order but can be omitted (Excel rebuilds it on open)
19. **External references** — `[1]Sheet1!A1` syntax with an external link table in `xl/externalLinks/`
20. **Encoding of special chars in sheet names** — Apostrophe-wrapped names with embedded apostrophes: `='Sheet''s Data'!A1`

---

## 5. Current Ascend Coverage Assessment

Based on the codebase exploration:

**Already implemented well:**
- Cell value types (number, string, boolean, date, error, empty, array)
- StyleRegistry (interned styles, ID-based references)
- Date system support (1900/1904 in FunctionEvalContext)
- Volatile function marking in FunctionDef
- Formula parser and evaluator
- 24 function modules covering math, stats, text, date, logical, lookup, financial, info, engineering, database, dynamic, convert
- Conformance test suite with `evalFormula()` harness
- Wildcard matching for COUNTIF/SUMIF criteria

**Gaps to investigate:**
- No XLSX test fixtures in `fixtures/` (all fixtures are JSON formula conformance)
- Shared formula expansion (verify in io-xlsx)
- Rich text round-trip (verify in io-xlsx)
- Number format string parsing completeness
- Conditional formatting evaluation
- `_xlfn.` prefix handling
- Dynamic array spill mechanics
- Defined name scope resolution in formula evaluation
- Shared string table edge cases (duplicates, inline strings)

---

## References

- [SheetJS Test Files](https://oss.sheetjs.com/test_files/)
- [Apache POI Test Data](https://github.com/apache/poi/tree/trunk/test-data/spreadsheet)
- [XlsxWriter Comparison Tests](https://github.com/jmcnamara/XlsxWriter/tree/master/xlsxwriter/test/comparison)
- [HyperFormula Excel Compatibility](https://hyperformula.handsontable.com/guide/compatibility-with-microsoft-excel.html)
- [ExcelNumberFormat Parser](https://github.com/andersnm/ExcelNumberFormat)
- [OOXML Validator](https://github.com/mikeebowen/OOXML-Validator)
- [ISO/IEC 29500 Schemas](https://github.com/sc34wg4/OOXMLSchemas)
- [MS-OI29500 Implementation Notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500)
- [MS-OE376 numFmt Spec](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/0e59abdb-7f4e-48fc-9b89-67832fa11789)
- [Excel 1900 Leap Year Bug](https://learn.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year)
- [Excel Recalculation Model](https://learn.microsoft.com/en-us/office/client-developer/excel/excel-recalculation)
- [Calamine](https://github.com/tafia/calamine)
- [XLParser](https://github.com/spreadsheetlab/XLParser)
- [POI Regression Test Corpus](https://github.com/centic9/poi-regression-test)
- [Acuity Excel Usage Statistics 2026](https://www.acuitytraining.co.uk/news-tips/excel-ai-data-statistics/)
- [12 Functions That Drive 97% of Financial Models](https://pps.financial/resources/the-12-excel-functions-that-drive-97-of-financial-models)
