# Formula binding role primitives

## Question

Can a read-only formula binding-role analyzer be folded into formula assist now, while still refusing to claim safe rename?

## Hypothesis

Yes. A stateless formula pass can classify local LET declarations/uses and structured-reference table/column roles from source spans. Workbook-scoped defined names and table ownership should remain unresolved until a workbook-context analyzer exists.

## External sources checked

- [Microsoft names in formulas](https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1) documents workbook and worksheet name scopes, including worksheet-local precedence over workbook names.
- [Microsoft named ranges guidance](https://learn.microsoft.com/office/vba/excel/Concepts/Cells-and-Ranges/refer-to-named-ranges) distinguishes worksheet-specific names from workbook-global names.
- [Microsoft structured references](https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e) documents table/column formulas and automatic updates after table or column renames.
- [MS-OI29500 structured references](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/089fbdef-ed49-4a14-9509-794c95651b17) defines table-name, column-name, and reserved keyword structure.
- [pygls rename example](https://pygls.readthedocs.io/en/latest/servers/examples/rename.html) shows `prepareRename` as an explicit guard before mutating rename operations.

## Why this matters to Ascend

Formula spans and reference lists help agents inspect formulas, but safe rename needs symbol roles. This fold-in gives clients enough read-only evidence to distinguish some declarations and uses without pretending that every name token is safe to rewrite.

## Probe/implementation

- Added `FormulaBindingRole` and `FormulaBindingRoleKind`.
- Added `formulaBindingRoles(formula)`.
- Added `bindings` to `FormulaAssistResult`.
- Classified:
  - `let-binding-declaration`
  - `let-binding-use`
  - `table-name-use`
  - `table-column-use`
  - `unresolved-name`
- Suppressed structured-reference body tokens from falling through as unresolved names.
- Kept workbook-defined names unresolved because stateless formula assist does not know workbook/sheet scope.

## Results

Focused validation passed:

```bash
bun test packages/sdk/src/formula-edit.test.ts
bun test apps/cli/src/cli.test.ts -t "formula assist"
bun test apps/api/src/server.test.ts -t "formula-assist"
bun test apps/mcp/src/index.test.ts -t "formula_assist"
```

Local probe:

| Formula | Binding roles |
| --- | --- |
| `=LET(total,SUM(A1:A3),total/3)+Budget+Sales[[#Totals],[Amount]]` | `total` declaration/use, unresolved `Budget`, table `Sales`, column `Amount` |
| `=SUM(Sales[@[Units]],Sales[Amount])` | table `Sales`, columns `Units` and `Amount` |

## Confidence

High for stateless binding-role evidence over LET and structured references. Medium-low for safe rename because defined names, table ownership, sheet names, external workbook qualifiers, and nested LET shadowing still require workbook/context-aware resolution.

## Fold-in decision

Promote to product/DX and correctness loops as a read-only formula intelligence primitive. Do not promote safe rename. The next fold-in should add workbook-context binding roles for defined names and table ownership, then a `prepareRename`-style rejection surface before any edit-producing rename action.

## Next question

Can workbook-context formula binding roles resolve defined names and table ownership without changing formula evaluation semantics or introducing unsafe rename edits?
