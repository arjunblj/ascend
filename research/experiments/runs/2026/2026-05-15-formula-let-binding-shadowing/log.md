# Formula LET Binding Shadowing

## Question

Can formula language-service primitives make a credible "formula intelligence" claim today, or do rename/code actions still need one more correctness fold-in before promotion?

## Hypothesis

Ascend can credibly claim read-only formula language-service primitives, but not safe rename. A local probe should find whether the current binding-role analyzer handles nested `LET` scopes well enough to avoid misleading clients.

## External sources checked

- Microsoft Formula AutoComplete documents Excel's formula editor as a language surface with functions, names, and structured-reference completions. https://support.microsoft.com/en-gb/office/use-formula-autocomplete-6d13daa5-e003-4431-abab-9edef51fae6b
- Microsoft structured references docs describe table names, column names, and automatic reference updates after table edits. https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- HyperFormula cell-reference docs distinguish cell, range, sheet-qualified, and named-expression references, reinforcing that formula tooling needs typed reference evidence. https://hyperformula.handsontable.com/guide/cell-references.html
- Monaco `CodeAction` docs separate code actions from raw text edits; edit-producing actions need a defensible target range and kind. https://microsoft.github.io/monaco-editor/typedoc/interfaces/languages.CodeAction.html
- LibreOffice named ranges docs show names have validity and scope rules, so unresolved name tokens are not safe rename targets by default. https://help.libreoffice.org/latest/en-GB/text/scalc/guide/value_with_name.html

## Why this matters to Ascend

Formula intelligence is product-shaped when agents and humans can inspect formulas, hover symbols, list references, see diagnostics, and request guarded edits. It becomes unsafe if binding roles are wrong, because a future rename/code-action surface could rewrite a declaration or a shadowed local symbol as if it were a workbook-level name.

## Probe/implementation

Probe:

```bash
bun -e "import { formulaBindingRoles } from './packages/sdk/src/index.ts'; console.log(JSON.stringify(formulaBindingRoles('=LET(x,1,LET(x,2,x)+x)'), null, 2))"
```

Before the fix, `=LET(x,1,LET(x,2,x)+x)` produced duplicate and incorrect evidence:

- outer `x` declaration at `5:6`
- inner `x` at `13:14` incorrectly reported as a use
- inner result `x` at `17:18` reported twice
- outer result `x` at `20:21` reported once

Fold-in:

- Added binding target spans (`bindingStart`, `bindingEnd`) to `FormulaBindingRole` uses.
- Reworked stateless `LET` role collection to walk nested LET scopes recursively.
- LET value expressions can see earlier outer declarations.
- Inner LET declarations shadow outer declarations for their result expression.
- Inner declarations are no longer classified as uses, and nested uses are no longer duplicated.
- Added regression coverage for shadowing and for nested LET value expressions that reference an outer binding.

## Results

After the fix, the same probe returns:

- outer declaration `x` at `5:6`
- inner declaration `x` at `13:14`
- inner result `x` at `17:18` bound to `13:14`
- outer result `x` at `20:21` bound to `5:6`

Validation:

- `bun test packages/sdk/src/formula-edit.test.ts -t "classifies formula binding roles"`
- `bun test packages/sdk/src/formula-edit.test.ts`
- `bun test apps/cli/src/cli.test.ts -t "formula assist"`
- `bun test apps/api/src/server.test.ts -t "formula-assist exposes diagnostics"`
- `bun test apps/mcp/src/index.test.ts -t "formula_assist"`
- `bunx biome check packages/sdk/src/formula-edit.ts packages/sdk/src/formula-edit.test.ts`
- `bunx tsc --build`
- `bun run test:changed`

## Confidence

High for stateless `LET` scope and shadowing evidence. Medium for the overall formula intelligence claim, because workbook-context binding roles still need defined-name scope, table ownership, sheet-name references, and external-workbook boundaries.

## Fold-in decision

Promote to correctness loop and product/DX loop as a read-only formula intelligence improvement. Do not promote safe rename/code actions yet.

## Next question

Can workbook-context formula binding roles resolve defined-name scope and table ownership strongly enough to support a `prepareRename` rejection surface without producing rename edits?
