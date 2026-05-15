# Formula Prepare Rename Boundary

## Question

Can workbook-context formula binding roles resolve defined-name scope and table ownership strongly enough to support a `prepareRename` rejection surface without producing rename edits?

## Hypothesis

Not fully. Ascend can safely expose a stateless `prepareRename`-style target for formula-local `LET` bindings after the shadowing fix, but workbook names, table names, table columns, sheet names, and cell/range references should be rejected until a workbook-context resolver exists.

## External sources checked

- Microsoft names-in-formulas docs: defined names can be workbook-scoped or worksheet-scoped. https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft named-range docs: worksheet-specific named ranges are not global across a workbook. https://learn.microsoft.com/office/vba/excel/Concepts/Cells-and-Ranges/refer-to-named-ranges
- Microsoft structured reference docs: table and column renames automatically update structured references across the workbook. https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- LSP 3.17 `prepareRename`: the response can be a range/placeholder, default behavior, or null. https://github.com/Microsoft/language-server-protocol/blob/gh-pages/_specifications/lsp/3.17/specification.md
- HyperFormula cell-reference docs: formula references include cells, ranges, sheet references, and named expressions, which are not all the same kind of rename target. https://hyperformula.handsontable.com/guide/cell-references.html

## Why this matters to Ascend

The product-shaped claim is "formula language-service primitives," not "string find/replace for formulas." A `prepareRename` boundary lets UI, CLI, API, and MCP clients ask whether the cursor is on a renameable formula symbol without applying edits or pretending workbook-wide rename safety exists.

## Probe/implementation

Local inspection showed:

- Defined-name scope exists in the core workbook model.
- Tables exist on sheets and table rename operations exist.
- Stateless formula assist currently sees `unresolved-name`, `table-name-use`, and `table-column-use`, but does not know workbook/sheet/table ownership.

Fold-in:

- Added `formulaPrepareRename(formula, cursor)`.
- Added `renameTarget` to `FormulaAssistResult` for cursor-aware API/CLI/MCP consumers.
- Exported `FormulaPrepareRenameResult`, range/reason types, and `formulaPrepareRename`.
- `LET` declarations and uses with resolved binding spans return `ok: true`, a placeholder, the declaration range, and all formula-local occurrence ranges.
- Unresolved names, table names, and table columns return `ok: false` with `workbook-context-required`.
- Cell/range/sheet/external references return `ok: false` with `reference-target-not-renameable`.
- No symbol at cursor returns `ok: false` with `no-symbol-at-cursor`.

Probe examples:

| Formula/cursor | Result |
| --- | --- |
| `=LET(x,1,LET(x,2,x)+x)` on inner result `x` | `ok: true`, declaration `13:14`, occurrences `13:14` and `17:18` |
| `=Budget+Sales[Amount]` on `Budget` | `ok: false`, `workbook-context-required` |
| `=Budget+Sales[Amount]` on `Sales` | `ok: false`, `workbook-context-required` |
| `=A1+B1` on `A1` | `ok: false`, `reference-target-not-renameable` |

## Results

This improves the formula intelligence claim without promoting unsafe rename. The surface now tells clients exactly where rename is currently safe (`LET` local bindings only) and where product/correctness work is still required.

Validation passed:

- `bun test packages/sdk/src/formula-edit.test.ts`
- `bun test apps/cli/src/cli.test.ts -t "formula assist"`
- `bun test apps/api/src/server.test.ts -t "formula-assist exposes diagnostics"`
- `bun test apps/mcp/src/index.test.ts -t "formula_assist"`
- `bunx biome check packages/sdk/src/formula-edit.ts packages/sdk/src/formula-edit.test.ts packages/sdk/src/index.ts`

Validation blocked by unrelated dirty journal worktree changes:

- `bunx tsc --build` fails in `packages/sdk/src/journal.ts` because a dirty journal change maps issue objects with string `code` values into a `MutationJournalIssue` callback.
- `bun run test:changed` fails in journal/API/MCP/agent workflow tests because dirty journal changes add `surface` and `reason` fields to issue objects while tests still expect the older exact shape.
- This cycle touched only `packages/sdk/src/formula-edit.ts`, `packages/sdk/src/formula-edit.test.ts`, and `packages/sdk/src/index.ts`.

## Confidence

High for stateless formula-local `LET` prepare-rename evidence. Medium-low for workbook-context rename, which remains deliberately blocked.

## Fold-in decision

Promote to product/DX and correctness loops as a guarded formula language-service primitive. Do not promote edit-producing rename until workbook-context roles resolve defined-name scope and table ownership.

## Next question

Can the agent-view claim be tightened with a fixture-backed token budget proof that compares compact agent views against raw sheet JSON for the same workbook?
