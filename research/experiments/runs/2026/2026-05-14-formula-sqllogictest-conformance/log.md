# Formula Sqllogictest Conformance

## Question

Can sqllogictest-style formula conformance make correctness claims easier to audit?

## Hypothesis

Yes. Ascend already has JSON formula conformance fixtures and a HyperFormula comparator. A line-oriented completion/validation format would make oracle provenance, skips, known divergences, and large generated cases easier to review.

## External sources checked

- SQLite sqllogictest documentation: https://www.sqlite.org/sqllogictest
- DuckDB sqllogictest introduction: https://duckdb.org/docs/current/dev/sqllogictest/intro.html
- DuckDB sqllogictest writing tests guide: https://duckdb.org/docs/current/dev/sqllogictest/writing_tests.html
- DuckDB sqllogictest result verification guide: https://duckdb.org/docs/current/dev/sqllogictest/result_verification.html

## Why this matters to Ascend's North Star

Correctness claims need auditability. A spreadsheet formula corpus that separates prototype cases from completed oracle results would let Ascend compare against Excel, HyperFormula, LibreOffice, or internal evaluators without hiding mismatches in test code.

## Boundary / what this will not change

This experiment does not replace JSON fixtures or change formula tests. It records a future harness shape.

## Probe implementation or evidence-gathering method

Command:

```bash
bun run fixtures/formulas/formula-hyperformula-compare.ts
```

External model: sqllogictest uses prototype scripts, completion mode, validation mode, labels, skips, and result hashing for large outputs.

Cycle 14 folded in a first reusable harness:

- Added `fixtures/formulas/formula-logictest.ts`.
- Added `emitFormulaLogicTest()` to convert current JSON conformance fixtures into completed line-oriented formula-logictest text.
- Added `parseFormulaLogicTest()` and `runFormulaLogicTest()` to validate completed formula-logictest text against Ascend's evaluator.
- Added `fixtures/formulas/formula-logictest.test.ts` covering conversion, parse round-trip, and validation against a smoke subset.

## Results

The current comparator reported:

- 763 scenarios.
- 0 unexpected mismatches.
- 35 known comparator divergences.
- 261 comparator skips.

This is already close to sqllogictest's discipline: it distinguishes match, known divergence, and skip. The current weakness is that the durable test artifact is split between JSON fixtures and TypeScript comparator policy. Reviewers cannot inspect one completed text artifact that says "this formula, this setup, this oracle, this expected result, these engines skipped".

Cycle 14 validation passed:

```bash
bun test fixtures/formulas/formula-logictest.test.ts
bunx biome check fixtures/formulas/formula-logictest.ts fixtures/formulas/formula-logictest.test.ts
bunx tsc --build
```

Candidate formula-logictest sketch:

```text
setup Sheet1!A1 number 1
setup Sheet1!A2 number 2
query value label=sum-basic
=SUM(A1:A2)
----
number 3

skipif hyperformula unsupported-function
onlyif excel
query error label=countblank-empty-string
=COUNTBLANK(A1:A3)
----
number 2
```

## Confidence

High that this would improve auditability. Medium on the exact format; it should be compatible with existing JSON fixtures rather than a wholesale rewrite.

## Fold-in recommendation:

- promote to correctness loop: yes, cycle 14 added the initial converter/runner beside existing fixtures
- promote to performance loop: no
- promote to product/DX loop: no
- promote to topic synthesis: yes
- archive as dead end: no

## Next best question

Can a converter emit a completed formula-logictest file from the current JSON fixtures and the Excel/HyperFormula oracle metadata without changing existing tests?
