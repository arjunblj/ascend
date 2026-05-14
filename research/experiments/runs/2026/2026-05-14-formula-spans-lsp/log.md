# Formula Spans LSP

## Question

Can formula AST spans unlock hover, diagnostics, completions, rename, and safe code actions like an LSP?

## Hypothesis

Ascend can expose a useful formula-language-service surface today from token spans and SDK helpers, but safe rename/code actions need AST nodes with source ranges and symbol binding roles. Token spans are enough for highlight, hover, active signature help, parse diagnostics, and simple reference insert/cycle actions. They are not enough to rename sheet/table/name/LET symbols safely.

## External sources checked

- [Language Server Protocol overview](https://microsoft.github.io/language-server-protocol/): LSP standardizes editor features such as completion, definition, references, and hover across tools.
- [VS Code Programmatic Language Features](https://code.visualstudio.com/api/language-extensions/programmatic-language-features): maps editor APIs to LSP methods including diagnostics, hover, completion, signature help, code actions, and rename.
- [VS Code Language Server Extension Guide](https://code.visualstudio.com/api/language-extensions/language-server-extension-guide): lists language features and highlights incremental document synchronization to avoid repeated full parsing.
- [Microsoft LAMBDA function documentation](https://support.microsoft.com/en-us/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67): named formula functions and parameters make formula binding more program-like.
- [HyperFormula cell references](https://hyperformula.handsontable.com/docs/guide/cell-references.html): references include named expressions, sheet-qualified refs, ranges, absolute refs, and copy/move semantics.
- [HyperFormula dependency graph](https://hyperformula.handsontable.com/docs/guide/dependency-graph.html): formula references drive graph edges and range nodes, reinforcing that references need exact ranges and roles.

## Why this matters to Ascend

Formula editing is a high-leverage human and agent DX surface. Exact spans make formulas inspectable: an agent can explain the symbol under the cursor, highlight precedents, offer safe reference insertion, show function signatures, and emit precise diagnostics. Safe rename and code actions would be a leapfrog feature over spreadsheet libraries that treat formulas as opaque strings.

## Probe/implementation

Inspected local implementation:

- `packages/formulas/src/tokens.ts` gives each token a `position`.
- `packages/formulas/src/ast.ts` has semantic node types but no `start`, `end`, or `rawText`.
- `packages/formulas/src/parser.ts` parses structured refs, sheet-qualified refs, sheet spans, workbook-qualified sheet spans, function calls, ranges, arrays, and reference operators, but node construction drops token ranges.
- `packages/sdk/src/formula-edit.ts` already exposes `formulaTokenRanges`, `referenceAtCursor`, `cycleFormulaReferenceMode`, `insertFormulaReference`, `formulaDiagnostics`, `formulaFunctionCompletions`, and `formulaFunctionSignatureHelp`.
- `packages/sdk/src/formula-edit.test.ts` covers reference lookup, diagnostics, completions, signature help, token ranges, and F4-style reference cycling.

Updated the ignored probe `research/experiments/runs/2026/2026-05-14-formula-spans-lsp/probes/formula-token-spans.ts`. It evaluates formulas with structured refs, nested `IF`, sheet spans, external workbook sheet spans, `XLOOKUP`, `LET`, and malformed structured refs. For each case it records parse status, token spans, reference-at-cursor results, signature help, completions, diagnostics, and F4-cycle behavior.

Validation commands:

```bash
bun run research/experiments/runs/2026/2026-05-14-formula-spans-lsp/probes/formula-token-spans.ts
bun test packages/sdk/src/formula-edit.test.ts
```

## Results

Useful capabilities already exist without production changes:

| Capability | Current evidence | Ready for fold-in? |
| --- | --- | --- |
| Syntax highlighting | `formulaTokenRanges` returns stable ranges and syntax classes | yes |
| Reference hover/selection | `referenceAtCursor` resolves cell, range, sheet, 3D sheet, workbook-qualified, structured, and spill refs | yes |
| Signature help | `formulaFunctionSignatureHelp` returns active function and parameter, including nested calls | yes |
| Function completion | `formulaFunctionCompletions` returns registry-backed signatures | yes |
| Diagnostics | `formulaDiagnostics` returns parse and malformed-reference spans | yes |
| Safe reference insertion | `insertFormulaReference` can replace the active reference or insert into empty slots | yes |
| F4 reference cycling | `cycleFormulaReferenceMode` rewrites only the target cell endpoint | yes for cell/range refs |
| Rename/code actions | token spans alone do not identify symbol bindings or AST roles | no |

Probe details:

- `=SUM(Table1[Amount])` resolved `Table1[Amount]` as a structured reference and showed `SUM(arg1, [arg2], ...)`.
- `=XLOOKUP("sku-7",Products[SKU],Products[Price])` identified both structured references and active parameters 1 and 2.
- `=LET(total,SUM(A2:A10),total/COUNT(A2:A10))` showed LET and nested function signature help, but `total` remained a plain `Name` token with no binding role.
- `='[Budget.xlsx]Jan:Mar'!$B$2:$C$9` parsed as a `sheetSpanRef`, and the SDK could select the full workbook-qualified 3D range.
- `=SUM(Table1[[#Totals],[Amount])` produced a structured-reference diagnostic at the bracketed span plus a parse error at EOF.
- Targeted validation passed: `bun test packages/sdk/src/formula-edit.test.ts` ran 15 tests with 0 failures.

## Confidence

High for promoting current SDK formula-edit helpers as product/DX language-service primitives. Medium for AST-span work: the parser can be instrumented, but the design must avoid disrupting evaluator ergonomics and existing tests.

## Fold-in decision

Promote to the product/DX loop now: expose formula language-service primitives in CLI/API/MCP for agents and UI.

Promote to the correctness loop for a scoped AST span design: add source ranges and binding-role tests before implementing rename/code actions.

Do not promote to the performance loop yet. Incremental parsing is relevant later, but correctness and API shape should come first.

## Next question

Can property-based operation testing prove inverse journal laws better than the current hand-authored operation fixtures?
