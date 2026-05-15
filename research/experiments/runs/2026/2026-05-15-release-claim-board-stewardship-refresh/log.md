# Release Claim Board Stewardship Refresh

## Question

Can Ascend stop promoting new surfaces for this block and turn the current experiment evidence into a release-claim board with explicit proof obligations, while keeping formula intelligence rejection-first and not implementing rename?

## Hypothesis

Yes. The highest-value next research output is a constrained board: allowed claim wording, missing proof by dimension, and owner loop. Formula intelligence should remain a primitive/rejection claim until workbook-context rename proof exists.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- LSP 3.17 `prepareRename`: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft LET function: https://support.microsoft.com/en-au/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
- Microsoft names in formulas: https://support.microsoft.com/en-gb/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft structured references: https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- openpyxl tutorial and preservation warnings: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/
- HyperFormula cell references: https://hyperformula.handsontable.com/guide/cell-references.html
- Apache Arrow documentation: https://arrow.apache.org/docs
- DuckDB Excel import: https://duckdb.org/docs/stable/guides/file_formats/excel_import

## Why this matters to Ascend

The research loop had started finding implementation surfaces faster than release claims could absorb them. A claim board keeps Ascend honest: product wording must map to fixture proof, benchmark proof, public API/CLI/MCP surface proof, validation gates, competitor contrast, and an explicit boundary.

## Probe/implementation

- Ran `git status --short --branch`; no tracked production diff needed finishing.
- Inspected formula assist/rename references in SDK, CLI, API, and MCP.
- Validated current formula guard behavior with `bun test packages/sdk/src/formula-edit.test.ts`.
- Updated `research/experiments/syntheses/2026-05-release-claim-board.md` so each claim row uses the required three columns and enumerates proof still missing by:
  - fixture;
  - benchmark;
  - API/CLI/MCP surface;
  - validation gate;
  - competitor contrast;
  - honest boundary.

## Results

The board now explicitly says what Ascend may claim today and what proof remains before stronger release copy:

- Safe unknown workbook opening remains the top product/performance handoff.
- Auditable package-part mutation remains the top correctness/product handoff.
- Token-bounded agent view, retained viewport patch history, formula language-service primitives, release proof bundle, formula oracle routing, and columnar sidecars remain proof backlog unless their owner loop closes the listed proof gaps.

Formula intelligence remains rejection-first:

- binding roles are not mutation authority;
- LET shadowing must be lexical and nearest-scope-wins;
- defined names require workbook/sheet scope resolution;
- table names and columns require workbook/table ownership resolution;
- external, sheet, cell, range, 3D, and spill refs are not local rename targets;
- `prepareRename` may only return success for proven formula-local LET evidence and must not apply edits.

## Confidence

High. This pass intentionally narrows claims and does not introduce production behavior. The formula guard tests pass, and the board points future implementation loops at proof packaging rather than new surfaces.

## Fold-in decision

Promote to topic synthesis only. Do not fold in production code. Hand off only the top one or two claims:

1. safe unknown workbook opening to product/performance;
2. auditable package-part mutation to correctness/product.

## Next question

Can the safe unknown workbook opening owner produce a release-environment proof report from existing surfaces and public fixtures without adding a new command, endpoint, or MCP tool?
