# Case-insensitive binding refs

## Question

Can Verify stop false-positive formula-binding integrity errors when imported formula metadata uses sheet-qualified references whose sheet-name casing differs from the workbook sheet name?

## Hypothesis

Yes. Sheet-qualified binding evidence should compare sheet names case-insensitively while preserving the original text for reporting. A targeted verifier fix can accept `data!A1:A2` as binding evidence for a workbook sheet named `Data` without weakening cell/range checks.

## External sources checked

- Microsoft worksheet rename rules: https://support.microsoft.com/en-gb/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
- Microsoft names-in-formulas documentation, including name case and workbook/worksheet scope: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft Open XML formulas overview: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- Microsoft worksheet references overview: https://learn.microsoft.com/nb-no/office/client-developer/excel/worksheet-references

## Why this matters to Ascend

Ascend's North Star depends on trustworthy workbook inspection and mutation evidence. Formula-binding verification should reject stale or impossible metadata, but it should not mark a workbook risky because a producer preserved a sheet-qualified reference with different casing. False positives weaken the credibility of auditable mutation and formula-intelligence proof.

## Probe/implementation

Inspected `packages/verify/src/checker.ts` and the formula-binding integrity tests. The in-flight fix added case-insensitive sheet-name comparison for formula-binding range containment, range equality, overlap checks, range keys, and spill-anchor checks.

Committed production change:

```bash
git show --stat --oneline a2860cbf
```

Commit:

```text
a2860cbf fix(verify): accept case-insensitive binding refs
```

Validation:

```bash
bun test packages/verify/src/verify.test.ts -t "case-insensitive sheet-qualified formula binding refs"
bun test packages/verify/src/verify.test.ts
bun test packages/core/src
bunx biome check packages/core/src/sparse-grid.ts packages/verify/src/checker.ts packages/verify/src/verify.test.ts
```

## Results

- The targeted verifier test accepts a workbook sheet named `Data` with binding metadata using `data!A1:A2`, `DATA!C1:C2`, and a `DATA!C1` spill anchor.
- Full `packages/verify/src/verify.test.ts` passed: 147 tests.
- `packages/core/src` passed: 115 tests.
- Biome passed for the touched production/test files.
- No formula rename surface was added.

## Confidence

High for the fixed false-positive class: the change is scoped to sheet-name equality inside binding-integrity checks, and validation covers the array/spill metadata path that triggered the risk. Medium for broader Excel producer behavior because this probe does not attempt to enumerate every external writer's sheet-name casing conventions.

## Fold-in decision

Promote to correctness loop and already folded in as `a2860cbf`. This is a verifier correctness fix, not a new formula language-service surface.

## Next question

Should formula-binding verification also normalize workbook-scoped defined-name casing, or would that blur the current rejection-first boundary for workbook-context formula intelligence?
