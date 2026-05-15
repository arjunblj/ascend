# Formula Assist Corpus Proof

Date: 2026-05-15

## Question

Can Ascend's formula language-service claim move from rejection-first surface snapshots to a corpus-backed proof that `formulaAssist` remains fast and refuses unsafe rename targets across realistic public formulas?

## Hypothesis

A tracked proof harness can run current formula-assist primitives over public formula workbooks plus explicit LET/table/name/external-reference edge cases, proving the allowed claim without implementing edit-producing rename.

## External Sources Checked

- LSP 3.17 `textDocument/prepareRename` defines a preflight request that can refuse invalid rename positions by returning no target: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft LET documents formula-local names scoped to `LET`: https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
- Microsoft structured references document table and column names in formulas and workbook-wide rename behavior: https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- HyperFormula documents dependency-graph/reference scope as a formula-engine benchmark contrast, not an XLSX-preserving edit guard: https://hyperformula.handsontable.com/docs/guide/dependency-graph.html

## Why This Matters To Ascend

Formula intelligence is a credible product claim only if it is useful to agents and humans without overclaiming unsafe mutation. Ascend can be stronger than generic spreadsheet libraries by exposing spans, hover, diagnostics, completions, code actions, and refusal reasons while explicitly keeping workbook-context renames out of scope until a correctness-owned planner exists.

## Probe/Implementation

Added `fixtures/benchmarks/formula-assist-proof.ts` and `fixtures/benchmarks/formula-assist-proof.test.ts`.

The harness:

- discovers formulas from public XLSX fixtures:
  - `fixtures/xlsx/poi/formula_stress_test.xlsx`
  - `fixtures/xlsx/poi/FormulaEvalTestData_Copy.xlsx`
  - `fixtures/xlsx/poi/evaluate_formula_with_structured_table_references.xlsx`
  - `fixtures/xlsx/closedxml/Misc_FormulasWithEvaluation.xlsx`
- runs `formulaAssist` over a bounded public sample and static rejection-first edge cases;
- reports parse/diagnostic counts, reference spans, binding roles, refusal counts, and optional latency percentiles;
- asserts no edit-producing rename boundary by treating `prepareRename` as classification only.

Validation command:

```bash
bun test fixtures/benchmarks/formula-assist-proof.test.ts
bun run fixtures/benchmarks/formula-assist-proof.ts --public-formula-limit 250
bunx biome check fixtures/benchmarks/formula-assist-proof.ts fixtures/benchmarks/formula-assist-proof.test.ts
```

## Results

Latest local proof run:

- Public formulas discovered: 1685
- Sampled formulas: 250
- Static rejection-first edge cases: 10
- Parse OK formulas: 260
- Diagnostic formulas: 0
- Reference spans: 506
- Binding roles: 19
- Longest sampled formula: 50 chars
- Prepare-rename OK targets: 3, all LET-local
- Prepare-rename refusals:
  - `no-symbol-at-cursor`: 40
  - `workbook-context-required`: 3
  - `reference-target-not-renameable`: 214
- Median assist latency: 0.0252 ms
- P95 assist latency: 0.0531 ms
- Max assist latency: 2.4206 ms

Edge cases covered:

| Case | Expected | Observed | Boundary |
| --- | --- | --- | --- |
| LET local binding | ok | ok | Formula-local only; no edits applied. |
| LET shadowed inner use | ok | ok | Nearest binding only. |
| LET shadowed outer use | ok | ok | Outer binding not confused with inner binding. |
| Defined name | `workbook-context-required` | `workbook-context-required` | Requires workbook scope and collision rules. |
| Table name | `workbook-context-required` | `workbook-context-required` | Requires workbook table ownership. |
| Table column | `workbook-context-required` | `workbook-context-required` | Requires table schema ownership. |
| External ref | `reference-target-not-renameable` | `reference-target-not-renameable` | Workbook/path operation, not formula-local rename. |
| 3D ref | `reference-target-not-renameable` | `reference-target-not-renameable` | Workbook operation. |
| Spill ref | `reference-target-not-renameable` | `reference-target-not-renameable` | Reference semantics. |
| Function token | `no-symbol-at-cursor` | `no-symbol-at-cursor` | Not a rename target. |

## Confidence

Medium-high for the narrowed claim: formula-assist primitives are corpus-backed and rejection-first. Confidence stays low for any safe workbook-context rename claim because this proof intentionally does not resolve workbook names, table schemas, chart references, data validations, conditional formats, or external workbooks.

## Fold-In Decision

Promote to topic synthesis and product/DX proof packaging, not to new production rename work.

Allowed claim wording can now say: Ascend exposes formula language-service primitives with corpus-backed latency evidence and rejection-first prepare-rename classification.

Still forbidden: safe rename, edit-producing rename, workbook-context rename, table/defined-name rename, or cross-workbook rewrite.

## Next Question

Should the correctness loop build a workbook-context symbol ownership proof for defined names and table columns, or should formula intelligence stay frozen until the top two release claims, safe unknown workbook opening and auditable package-part mutation, are packaged?
