# Formula Oracle Routing JSON

## Question

Can formula corpus mismatch records expose explicit oracle routes and counters without changing existing correctness thresholds?

## Hypothesis

Yes. The corpus runner already classifies cached-value differences into semantic mismatches, numeric drift, stale cached oracles, and volatile skips. A route layer can explain the next oracle for each record while leaving pass/fail thresholds intact.

## External sources checked

- [HyperFormula known limitations](https://hyperformula.handsontable.com/docs/guide/known-limitations.html): documents cases such as multiple workbooks, 3D references, dynamic arrays, and structured references that should not be routed blindly to HyperFormula.
- [LibreOffice Recalculate Hard](https://help.libreoffice.org/latest/en-US/text/scalc/01/recalculate_hard.html?DbPAR=CALC): confirms LibreOffice can force recalculation of all formula cells, including non-volatile functions.
- [Apache POI formula evaluation](https://poi.apache.org/components/spreadsheet/eval.html): documents cached formula values, formula evaluator caches, and setup for referenced workbooks.
- [Excel calculation performance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance): documents smart recalculation, dependency tracking, full calculation, and full dependency rebuild.
- [Excel Application.CalculateFullRebuild](https://learn.microsoft.com/office/vba/api/Excel.Application.CalculateFullRebuild): documents full calculation plus dependency rebuild for Excel workbooks.

## Why this matters to Ascend

Formula conformance claims are only credible if every mismatch says what kind of evidence should close it. Cached values are useful, but they are not enough for volatile formulas, stale upstream caches, external links, dynamic arrays, structured references, date-system differences, and pure scalar semantic disagreements.

Route fields make the corpus JSON more auditable for humans and agents: accepted mismatch, Excel full rebuild, LibreOffice hard recalc, HyperFormula compatibility, static cached values, or manual triage.

## Probe/implementation

Inspected local implementation:

- `fixtures/benchmarks/formula-corpus-correctness.ts` already emits JSON payloads with mismatch classifications, volatile skips, accepted/unaccepted counts, stale-cache allowlist records, downstream propagation, and assertion gates.
- `fixtures/benchmarks/formula-corpus-correctness.test.ts` already covers semantic mismatches, numeric drift, volatile skips, downstream propagation, no-cache skips, CLI gates, and TypeScript manifests.

Folded in a scoped production harness change:

- Added `FormulaOracleRoute` values: `accepted-mismatch`, `cached-values`, `excel-full-rebuild`, `hyperformula-compat`, `libreoffice-hard-recalc`, and `manual-triage`.
- Added `oracleRoute`, `oracleReason`, and `oracleArtifact` fields to mismatch and volatile-skip JSON records.
- Added per-workbook and suite-level `oracleRouteCounts`.
- Routed numeric drift, stale cache, and volatile skips to `accepted-mismatch`.
- Routed external workbook references, structured references, Excel-specific/dynamic-array functions, and date-system-sensitive formulas to `excel-full-rebuild`.
- Routed LibreOffice-origin semantic mismatches to `libreoffice-hard-recalc`.
- Routed remaining semantic scalar mismatches to `hyperformula-compat`.
- Kept all existing threshold gates unchanged.

## Results

Validation passed:

```bash
bun test fixtures/benchmarks/formula-corpus-correctness.test.ts -t "classifies semantic, numeric drift, and volatile cached-value oracle skips"
bun test fixtures/benchmarks/formula-corpus-correctness.test.ts
bunx biome check fixtures/benchmarks/formula-corpus-correctness.ts fixtures/benchmarks/formula-corpus-correctness.test.ts
bunx tsc --build
```

The focused classification test now verifies:

| Case | Route |
| --- | --- |
| numeric drift | `accepted-mismatch` |
| pure scalar semantic mismatch | `hyperformula-compat` |
| volatile skip | `accepted-mismatch` |
| date-system-sensitive semantic mismatch | `excel-full-rebuild` |

Full formula corpus correctness tests passed: 13 tests, 0 failures.

## Confidence

High for the JSON schema extension and counters. Medium for the exact routing heuristics; they are intentionally conservative and should be tuned as real oracle runners land.

## Fold-in decision

Folded into the correctness harness.

Next fold-in should add either a real Excel/LibreOffice/HyperFormula runner behind these route fields or a compact release-proof summary that reports route counts alongside formula corpus results.

## Next question

Can package write plans expose a small action taxonomy over existing write evidence: passthrough, regenerate, add, drop, and error?
