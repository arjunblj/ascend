# Table Column Validation Boundary

## Question

Should `setTableColumn` reject invalid public column selectors and metadata values before mutating table columns or cells?

## Hypothesis

Yes. Table column edits affect headers, calculated-column formulas, totals rows, structured reference rewrites, and formula binding materialization. Invalid selectors, names, formulas, or totals metadata should fail before mutation and should be classified by journal exactness.

## External sources checked

- Microsoft Excel table overview: https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c
- Microsoft Excel calculated columns: https://support.microsoft.com/en-us/office/use-calculated-columns-in-an-excel-table-873fbac6-7110-4300-8f6f-aafa2ea11ce8
- OOXML table columns reference: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_tableColumns_topic_ID0EVSU5.html

## Why this matters to Ascend

The auditable mutation claim depends on table metadata edits being explainable and reversible when public operations represent them. `setTableColumn` is especially sensitive because it can rewrite structured references and materialize formula bindings. Rejecting invalid public metadata before mutation keeps that path trustworthy.

## Probe/implementation

Implemented:

- engine validation for `setTableColumn.column`, `newName`, `formula`, `totalsRowFormula`, `totalsRowFunction`, and `totalsRowLabel`.
- engine regression coverage proving invalid inputs reject without changing table metadata or nearby cells.
- SDK journal exactness classification for the same invalid values as `UNSUPPORTED_VALUE`.

## Results

Focused validation:

- `bun test packages/engine/src/operations.test.ts -t "setTableColumn rejects invalid"`
- `bun test packages/sdk/src/journal-exactness.test.ts -t "classifies missing required metadata updates as unsupported values"`
- `bunx biome check packages/engine/src/operations/table-ops.ts packages/engine/src/operations.test.ts packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts`
- `bunx tsc --build`

## Confidence

High. The patch is narrow, mirrors existing journal value classification, and covers table/cell non-mutation for rejected inputs.

## Fold-in decision

Promote to correctness loop as auditable mutation hygiene. This does not add a new table surface or promote stronger release wording.

## Next question

Run `bun run test:changed` after the remaining unrelated RC gate and IO reader edits either land or clear, so table metadata validation is confirmed in a cleaner worktree.
