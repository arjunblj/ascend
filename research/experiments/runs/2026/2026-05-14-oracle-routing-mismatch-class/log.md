# Oracle Routing By Mismatch Class

## Question

Can Excel, LibreOffice, HyperFormula, static cached values, and manual review be routed automatically by mismatch class instead of treating formula conformance as one monolithic oracle?

## Hypothesis

Yes. Ascend already classifies cached-value mismatches into semantic, numeric drift, stale oracle, and volatile/downstream skip paths. Adding an explicit oracle-routing layer should make formula correctness claims more auditable: each mismatch can name the next oracle, the reason it was chosen, and the artifact required to close the case.

## External sources checked

- [HyperFormula compatibility with Microsoft Excel](https://hyperformula.handsontable.com/docs/guide/compatibility-with-microsoft-excel.html): documents Excel compatibility knobs and function coverage limits.
- [HyperFormula known limitations](https://hyperformula.handsontable.com/docs/guide/known-limitations.html): confirms gaps that matter for routing, including multiple workbooks, 3D references, dynamic arrays, structured references, and some named-expression behavior.
- [Apache POI formula evaluation](https://poi.apache.org/components/spreadsheet/eval.html): explains cached formula values, recalculation, external workbook references, evaluator cache behavior, and cross-workbook setup.
- [LibreOffice Recalculate Hard](https://help.libreoffice.org/latest/sid/text/scalc/01/recalculate_hard.html): documents a hard recalc path for Calc formula cells, including non-volatile functions.
- [Office Scripts ExcelScript.Application](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.application?view=office-scripts): documents Excel application recalculation, calculation mode/state, calculation engine version, and culture/separator metadata.

## Why this matters to Ascend

Formula correctness will always involve more than one truth source. Excel is authoritative for Excel semantics, but it is not always available in CI. LibreOffice is scriptable and useful for open-source fixtures, but it is not Excel. HyperFormula is useful for fast formula/dependency cross-checks, but its own documentation says some Excel features are outside its current model. Static cached values preserve real workbook history, but stale caches and volatile formulas need explicit treatment.

An oracle router would let Ascend report formula evidence as a chain of decisions instead of a single pass/fail count. That is a stronger story for agents: each mismatch becomes a routed case with a reproducible proof artifact.

## Probe/implementation

Inspected local implementation:

- `fixtures/benchmarks/formula-corpus-correctness.ts` has a `cached-values` oracle mode, `semantic | numeric-drift | stale-oracle` mismatch classes, a volatile-function skip path, known stale-oracle records, downstream propagation for volatile/numeric/stale precedents, JSON output, and threshold assertions.
- `fixtures/xlsx/libreoffice-fixtures.test.ts` already asserts cached formula parity across LibreOffice formula fixtures with accepted mismatch counts: 34 workbooks, 418 formulas, 402 compared formulas, 23 accepted mismatches, and 0 semantic mismatches.
- `fixtures/formulas/excel-ground-truth.ts` builds manual Excel ground-truth workbooks and verifies recalculated cached values, but it is not yet an automatic oracle route.

Added ignored probe `research/experiments/runs/2026/2026-05-14-oracle-routing-mismatch-class/probes/oracle-routing-candidate.ts`. The probe models ten mismatch classes and routes each to one of:

- `accepted-mismatch`
- `cached-values`
- `excel-full-rebuild`
- `hyperformula-compat`
- `libreoffice-hard-recalc`
- `manual-triage`

Validation command:

```bash
bun run research/experiments/runs/2026/2026-05-14-oracle-routing-mismatch-class/probes/oracle-routing-candidate.ts
```

## Results

The probe routed all expected synthetic cases with no failures:

| Mismatch class | Route | Artifact |
| --- | --- | --- |
| `numeric-drift` | `accepted-mismatch` | accepted mismatch record with cache/calculated hashes |
| `volatile` | `accepted-mismatch` | volatile skip record and downstream propagation |
| `stale-cache` | `accepted-mismatch` | known stale-oracle record with reason |
| `unsupported-function` | `excel-full-rebuild` | Excel full-rebuild workbook plus engine metadata |
| `external-reference` | `excel-full-rebuild` | Excel or POI-style linked workbook bundle |
| `dynamic-array` | `excel-full-rebuild` | Excel spill-aware cache diff |
| `structured-reference` | `excel-full-rebuild` | Excel table-aware cache diff |
| `date-system` | `excel-full-rebuild` | Excel cache diff with date system and culture metadata |
| `semantic` | `hyperformula-compat` or `libreoffice-hard-recalc` | formula transcript or LibreOffice hard-recalc cache diff |
| `oracle-error` | `manual-triage` | oracle failure bundle |

This exposed a clear fold-in shape: keep the existing cached-value corpus runner, but add a second layer that converts mismatch records into routed cases. The router should not replace current thresholds; it should explain what evidence is still missing.

## Confidence

Medium-high. The routing taxonomy is strongly supported by external tool behavior and by Ascend's existing mismatch classifications. The probe is synthetic, so confidence is in the decision model rather than in a production harness.

## Fold-in decision

Promote to correctness loop.

Recommended fold-in:

1. Add `oracleRoute`, `oracleArtifact`, and `oracleReason` fields to formula corpus mismatch JSON output.
2. Keep current threshold assertions intact.
3. Add route-specific counters so release proof bundles can say how many mismatches are accepted, Excel-routed, LibreOffice-routed, HyperFormula-routed, static-golden, or manual.
4. Add an optional future `--emit-oracle-routes` mode before wiring any external Excel or LibreOffice runner into CI.

Do not fold into production in this research loop. The next correctness loop should make the JSON schema change with tests.

## Next question

Can Arrow/DuckDB-style columnar sidecars accelerate table and range scans without replacing workbook semantics?
