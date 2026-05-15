# Copy Sheet Table Retargeting

## Question

Can `copySheet` preserve table semantics by assigning workbook-unique copied table identities and rewriting copied structured references?

## Hypothesis

Yes. A copied sheet should not reuse the original table identity/name in a workbook where table names are global. Copied formulas, sheet metadata formulas, and sheet-scoped defined names that refer to copied tables should retarget to the copied table name. QueryTable-backed tables should fail closed because external table bindings cannot be safely duplicated by public operations yet.

## External sources checked

- Microsoft Support on Excel structured references: https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Microsoft Support on names in formulas and table name scope: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft Open XML SDK `TableParts` reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tableparts?view=openxml-2.20.0
- OfficeDev Open XML table guide: https://github.com/OfficeDev/open-xml-docs/blob/main/docs/spreadsheet/working-with-tables.md
- Microsoft Open XML SDK `TableDefinitionPart` reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.tabledefinitionpart?view=openxml-3.0.1
- OOXML query table part reference: https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Query_topic_ID0EQBBM.html

## Why this matters to Ascend

Copying a sheet is a high-trust mutation. If a copied sheet reuses table names or leaves formulas pointing at the source table, Ascend's auditable mutation and formula intelligence claims both weaken. The safe boundary is to retarget copied table references when the table is local and reject queryTable-backed copies until external binding semantics are owned.

## Probe/implementation

Implemented:

- copied tables get new `id`, `sheetId`, and workbook-unique names.
- copied structured references in cell formulas, sheet metadata formulas, and copied sheet-scoped defined names are retargeted to the copied table name.
- queryTable-backed source tables reject before mutation with explicit error details.
- table-rename helpers are exported from the formula rewrite module for copySheet reuse.
- prepared-plan copySheet now proves the written and reopened XLSX keeps copied table names, OOXML ids, table part paths, and sheet table relationships distinct by generating new table part paths after existing/preserved `xl/tables/tableN.xml` entries.

## Results

Focused validation:

- `bun test packages/engine/src/operations.test.ts -t "copySheet assigns workbook-unique copied table identities"`
- `bun test packages/engine/src/operations.test.ts -t "copySheet rejects queryTable-backed tables"`
- `bun test packages/sdk/src/agent-workflow.test.ts -t "prepared copySheet commits reopen workbook-unique table identities"`
- `bun test packages/sdk/src/agent-workflow.test.ts -t "table"`
- `bunx biome check packages/engine/src/operations/sheet-ops.ts packages/engine/src/operations.test.ts packages/engine/src/structural/formula-rewrite.ts`
- `bunx tsc --build`

## Confidence

Medium-high. The tests cover copied table identity, formula retargeting, metadata formula retargeting, sheet-scoped defined-name retargeting, and queryTable rejection. More public workbook fixture coverage would increase confidence for complex structured references.

## Fold-in decision

Promote to correctness loop. This is a workbook mutation correctness fix, not a new formula rename surface.

## Next question

Can real public table-heavy workbook fixtures be added to the release proof bundle after owner fixture policy settles?
