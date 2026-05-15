# Formula assist reference span list

## Question

Can formula language-service primitives move one step closer to a product claim without implying safe rename or full binding resolution?

## Hypothesis

Yes. `formulaAssist` already exposes tokens, diagnostics, hover, active reference, completions, signature help, insertion, and reference cycling. Adding a full non-overlapping reference span list gives agents and UI clients a stable primitive for highlighting, explain, and preview workflows while keeping safe rename out of scope.

## External sources checked

- [HyperFormula key concepts](https://hyperformula.handsontable.com/guide/key-concepts.html) describe a parser-backed AST and dependency graph, setting the bar for mature formula analysis.
- [Microsoft LSP RenameParams](https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.languageserver.protocol.renameparams?view=visualstudiosdk-2022) reinforces that rename is a position-sensitive language-server operation, not just string replacement.
- [VS Code programmatic language features](https://code.visualstudio.com/api/language-extensions/programmatic-language-features) separates hover, completion, signature help, code actions, and rename as distinct capabilities.
- [Microsoft structured references overview](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/089fbdef-ed49-4a14-9509-794c95651b17) shows why spreadsheet references need syntax-aware spans for table references and reserved keywords.

## Why this matters to Ascend

Formula intelligence is a high-leverage agent and human editing surface. A complete reference-span list lets clients explain or highlight all referenced regions without reparsing formula text. It is also an honest stepping stone toward safe code actions because the output is explicit about reference kinds while still not claiming binding roles.

## Probe/implementation

- Inspected `packages/sdk/src/formula-edit.ts` and existing formula assist tests.
- Added exported `formulaReferenceRanges(formula)`.
- Added `references` to `FormulaAssistResult`.
- Reused the existing reference collector but fixed range advancement so a range endpoint is not also emitted as a duplicate nested cell reference.
- Exported `formulaReferenceRanges` from the SDK index.
- Added focused tests for formula assist reference output and mixed reference lists.
- Ran a local probe over:
  - `=SUM(A1:B2,Sales[Amount],A1#)`
  - `=SUM(Sheet1:Sheet3!A1,"A1")`
  - `=LET(x,A1,x+B1)`

## Results

Focused validation passed:

```bash
bun test packages/sdk/src/formula-edit.test.ts
bun test apps/cli/src/cli.test.ts -t "formula assist"
bun test apps/api/src/server.test.ts -t "formula-assist"
bun test apps/mcp/src/index.test.ts -t "formula_assist"
```

Probe output showed:

- `A1:B2`, `Sales[Amount]`, and `A1#` as non-overlapping `range`, `structured`, and `spill` references.
- `Sheet1:Sheet3!A1` as a `sheet-3d-cell` reference.
- `LET(x,A1,x+B1)` reports `A1` and `B1` cell references, but does not classify `x` as a LET binding.

## Confidence

High for formula highlighting/explain primitives. Medium-low for safe code actions because LET names, table column binding, defined names, external workbook bindings, and sheet/table rename safety still need binding-role analysis beyond token spans.

## Fold-in decision

Promote to product/DX and correctness loops as a small formula-assist primitive. Do not promote safe rename yet. The next formula loop should design binding roles before any rename/code-action claim.

## Next question

Can formula assist expose binding-role evidence for LET names, defined names, table columns, and sheet-qualified references without changing formula evaluation semantics?
