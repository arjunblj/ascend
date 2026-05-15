# Formula binding role boundary

## Question

Can formula assist expose binding-role evidence for LET names, defined names, table columns, and sheet-qualified references without changing formula evaluation semantics?

## Hypothesis

Not safely yet. Ascend has useful token/reference spans and workbook formula metadata, but safe rename requires distinguishing declarations from uses and resolving workbook context such as defined-name scope, table ownership, LET bindings, and structured-reference column bindings.

## External sources checked

- [Microsoft names in formulas](https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1) documents workbook and worksheet name scopes, including local worksheet precedence over workbook names.
- [Microsoft named ranges VBA guidance](https://learn.microsoft.com/office/vba/excel/Concepts/Cells-and-Ranges/refer-to-named-ranges) reinforces worksheet-specific named ranges versus workbook-global names.
- [Microsoft structured references](https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e) documents table/column references and automatic update behavior when tables or columns are renamed.
- [MS-OI29500 structured references](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/089fbdef-ed49-4a14-9509-794c95651b17) defines table-name and reserved-keyword structure for table references.
- [pygls rename example](https://pygls.readthedocs.io/en/latest/servers/examples/rename.html) shows `prepareRename` as a guard step before rename, matching Ascend's need to reject unsafe formula rename locations.

## Why this matters to Ascend

Formula language-service primitives are valuable today, but "safe rename" is a stronger claim. If Ascend cannot prove whether a token is a declaration, a local LET use, a workbook-defined name, a sheet-scoped name, a table name, or a table column, then rename/code actions can silently rewrite the wrong symbol.

## Probe/implementation

- Inspected formula assist, formula metadata, defined-name inventory, structured-reference parsing, and rename-related operation tests.
- Ran a local probe comparing stateless `formulaAssist` and workbook-backed `wb.formula()` on formulas containing LET names, a defined name, cell references, and a structured reference.
- No production changes were made.

Probe summary:

```json
{
  "assistReferences": ["A1", "B1", "Sales[Amount]"],
  "assistNameTokens": ["x", "x", "Budget", "Sales", "[Amount]"],
  "formulaMetadataReferences": ["Budget", "Sales[Amount]", "x", "B1", "x"]
}
```

## Results

- Stateless `formulaAssist` now cleanly reports non-overlapping cell/range/structured/spill spans, but intentionally does not resolve name tokens.
- Workbook formula metadata reports `Budget` and LET `x` as name-like references, but does not classify declaration versus use.
- Structured references include table and column syntax, but a safe rename needs workbook/table ownership and column-binding validation.
- Defined names already have workbook/sheet scope in workbook inventory; formula assist does not consume that context.

## Confidence

High that safe rename should remain unpromoted. Medium that a scoped next implementation can add read-only binding-role evidence without changing evaluation or mutation semantics.

## Fold-in decision

Archive safe rename as not ready. Promote a narrow correctness/product design for read-only binding-role evidence:

- `let-binding-declaration`
- `let-binding-use`
- `defined-name-use`
- `table-name-use`
- `table-column-use`
- `sheet-name-qualifier`
- `external-workbook-qualifier`
- `unresolved-name`

Do not implement rename/code actions until those roles are fixture-backed.

## Next question

Can a read-only formula binding-role analyzer be built over parsed AST plus workbook context, with tests for LET, defined names, structured refs, external refs, and sheet spans?
