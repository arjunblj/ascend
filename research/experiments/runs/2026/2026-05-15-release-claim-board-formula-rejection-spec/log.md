# Release Claim Board And Formula Rejection Spec

## Question

Can Ascend pause surface promotion and steward release claims by defining what product wording is allowed today, what proof is missing, and which loop owns each next proof, while explicitly refusing formula rename beyond formula-local guard evidence?

## Hypothesis

Yes. The strongest research move now is to constrain future implementation: finish any already in-flight bug fix, then turn the claim ladder into a release board and make formula intelligence rejection-first.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl tutorial: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/
- LSP 3.17 specification: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft names in formulas: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft structured references: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- HyperFormula cell references: https://hyperformula.handsontable.com/guide/cell-references.html
- Apache Arrow columnar format: https://arrow.apache.org/docs/format/Columnar.html
- DuckDB Excel import: https://duckdb.org/docs/stable/guides/file_formats/excel_import

## Why this matters to Ascend

Ascend's North Star needs credible product-shaped claims, not scattered surfaces. A release claim board makes the next loops accountable to proof: fixture, benchmark, API/CLI/MCP shape, validation gate, competitor contrast, and honest boundary. The formula spec prevents the formula intelligence claim from slipping into unsafe rename.

## Probe/implementation

- Ran `git status --short --branch` and found an in-flight mutation-journal structured issue contract.
- Finished and committed that bug/fold-in checkpoint as `b4ce83e7 feat(sdk): stabilize journal issue schema`.
- Validated the checkpoint with:
  - `bun test packages/sdk/src/journal-exactness.test.ts packages/sdk/src/interactive-contract.test.ts`
  - `bun test packages/sdk/src/journal-compatibility.test.ts`
  - `bun test packages/sdk/src/agent-workflow.test.ts -t "loss audit"`
  - `bun test apps/cli/src/cli.test.ts -t "journal"`
  - `bun test apps/api/src/server.test.ts -t "journal"`
  - `bun test apps/mcp/src/index.test.ts -t "journal"`
  - `bunx biome check ...`
  - `bunx tsc --build`
  - `bun run test:changed`
- Inspected current `formulaBindingRoles` and `formulaPrepareRename` behavior in `packages/sdk/src/formula-edit.ts` and tests.
- Added `research/experiments/syntheses/2026-05-release-claim-board.md`.

## Results

The synthesis defines a three-column release claim board:

- claim wording allowed today;
- proof still missing;
- owner loop.

Top handoffs only:

1. Safe unknown workbook opening to product/performance.
2. Auditable package-part mutation to correctness/product.

Formula intelligence is explicitly scoped to primitives and rejection:

- binding roles are evidence, not mutation authority;
- `LET` shadowing must be lexical and nearest-scope-wins;
- defined names require workbook/sheet scope resolution;
- table names and columns require workbook/table ownership resolution;
- cell, range, sheet, 3D, spill, and external references are not local rename targets;
- `prepareRename` may only return `ok: true` for proven formula-local `LET` bindings and must not apply edits.

## Confidence

High for the claim board as a release stewardship artifact. It is grounded in current Ascend surfaces and primary references, and it narrows rather than expands implementation. Medium for exact ordering below the top two claims; token-bounded agent view and retained viewport patch history are both credible but need product proof packaging.

## Fold-in decision

Promote to topic synthesis only. Do not add new production surfaces in this block. Hand off only the top two proof loops.

## Next question

Can the safe unknown workbook opening proof bundle be generated from existing open-plan surfaces over public fixtures without adding a new command or API shape?
