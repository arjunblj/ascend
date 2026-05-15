# Product claim map

## Question

Which product-shaped claims can Ascend credibly make today, which need exactly one more fold-in, and which should stay speculative until stronger proof exists?

## Hypothesis

The highest-value research output right now is not another narrow production surface. It is a ranked claim map that constrains implementation loops to the next claims Ascend can actually prove: safe unknown workbook opening and auditable package-part mutation.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft ProtectedViewWindow API: https://learn.microsoft.com/cs-cz/office/vba/api/excel.protectedviewwindow
- Microsoft Excel digital signatures and code signing: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing
- openpyxl tutorial and preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/
- HyperFormula key concepts: https://hyperformula.handsontable.com/guide/key-concepts.html
- HyperFormula dependency graph: https://hyperformula.handsontable.com/guide/dependency-graph.html
- HyperFormula built-in functions: https://hyperformula.handsontable.com/docs/guide/built-in-functions.html
- Microsoft structured references: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- MS-OI29500 structured references: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/089fbdef-ed49-4a14-9509-794c95651b17
- Univer MCP guide: https://docs.univer.ai/guides/sheets/getting-started/mcp
- Univer MCP features: https://docs.univer.ai/guides/sheets/features/mcp
- DuckDB Excel extension: https://duckdb.org/docs/lts/core_extensions/excel.html
- Apache Arrow columnar format: https://arrow.apache.org/docs/format/Columnar.html

## Why this matters to Ascend

Ascend's North Star is SOTA OSS spreadsheet infrastructure for agents and humans. That requires proof-shaped product claims, not a pile of individually useful surfaces. A claim ladder lets the correctness, performance, and product loops converge on the smallest next proof instead of continuing to promote every nearby experiment.

## Probe/implementation

- Inspected `research/experiments/index.md` and prior syntheses to inventory completed experiments and existing fold-in decisions.
- Inspected SDK/CLI/API/MCP code by search for open-plan, package action proof, token-budgeted agent views, retained viewport tokens, and formula assist/binding roles.
- Rebuilt `research/experiments/syntheses/2026-05-claim-ladder.md` into a product claim map with:
  - ranked claim ladder;
  - proof requirements per claim: fixture, benchmark, API/CLI/MCP surface, validation gate, competitor contrast, and honest boundary;
  - top two handoffs only;
  - next-loop prompts for correctness, performance, and product;
  - do-not-promote list.

## Results

The claim map ranks:

1. Safe unknown workbook opening: credible today, needs proof packaging.
2. Auditable package-part mutation: needs one stable proof-schema fold-in.
3. Token-bounded agent view: credible today, needs product packaging.
4. Retained viewport patch history: credible in SDK, one product loop away.
5. Formula language-service primitives: useful primitives, not safe rename.
6. Release proof bundle: valuable wrapper, should wait for ranks 1 and 2.
7. Formula conformance/oracle routing: correctness backlog.
8. Columnar scan sidecars: speculative product claim.

Only the top two were handed off to implementation loops. The synthesis explicitly blocks promotion of columnar sidecars, safe formula rename, signed provenance, universal formula compatibility, collaborative sync, and private corpus claims.

## Confidence

High for the ranking as of this cycle because it is grounded in current local surfaces and primary external references. Medium for exact ordering between token-bounded agent view and retained viewport patch history; both are credible, but neither is as differentiating as safe opening or package-part mutation proof.

## Fold-in decision

Promote to topic synthesis and hand off only:

- safe unknown workbook opening proof bundle to product/performance loops;
- auditable package-part mutation proof schema to correctness/product loops.

Do not fold production code in this cycle.

## Next question

Can the safe unknown workbook opening proof bundle be generated from existing SDK/CLI/API/MCP surfaces over public fixtures, with latency evidence and no new product surface?
