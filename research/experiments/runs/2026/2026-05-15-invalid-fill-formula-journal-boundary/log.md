# Invalid Fill Formula Journal Boundary

## Question

Can Ascend's inverse journal claim exact rollback for `fillFormula` when the formula text is rejected by the engine?

## Hypothesis

No. `fillFormula` validates and parses the formula before applying the range edit. If parsing fails, the journal should classify the attempted edit as unsupported formula value evidence instead of emitting exact inverse cell operations.

## External sources checked

- Microsoft Excel formula overview: https://support.microsoft.com/en-us/office/overview-of-formulas-34519a4e-1e8d-4f4b-84d4-d642c4f63263
- Microsoft fill formulas into adjacent cells: https://support.microsoft.com/en-us/office/fill-a-formula-down-into-adjacent-cells-76d462cb-697c-4dcb-a0b4-df4d976abdc1
- HyperFormula parser and grammar documentation: https://hyperformula.handsontable.com/guide/basic-concepts.html

## Why this matters to Ascend

Formula intelligence and auditable mutation claims both depend on rejection-first behavior. A rejected formula edit must not leave proof artifacts that imply exact rollback, especially when agents use journals to decide whether a mutation plan is safe to commit.

## Probe/implementation

- Ran a local invalid-operation probe comparing engine apply behavior with journal exactness for multiple candidate operations.
- Found `fillFormula` with `formula='=1+'` returned engine `VALIDATION_ERROR` but journal exactness still reported `exact=true`.
- Folded in a small SDK fix in `packages/sdk/src/journal.ts`: invalid `fillFormula` text now emits `UNSUPPORTED_VALUE` on the `formulas` surface with reason `value-unsupported` and no inverse operations.
- Added `value-unsupported` to the formula exactness matrix and a regression test in `packages/sdk/src/journal-exactness.test.ts`.

## Results

Focused probe after the fix:

```json
{
  "applyOk": false,
  "applyCode": "VALIDATION_ERROR",
  "exact": false,
  "issueCount": 1,
  "surfaces": ["formulas"],
  "reasons": ["value-unsupported"]
}
```

Validation:

```bash
bun test packages/sdk/src/journal-exactness.test.ts
bunx biome check packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts
bunx tsc --build
```

All passed.

## Confidence

High for `fillFormula` invalid syntax. Medium for all formula edit rejection classes; `setFormula` currently accepts raw formula text without the same parse gate, so it was not changed in this pass.

## Fold-in decision

Promote to correctness loop as a small production fix. This is journal proof hygiene only; it does not promote new formula rename, edit action, or language-service surface area.

## Next question

Should `setFormula` and `fillFormula` intentionally differ on parse-time validation, or should formula edit validation be made consistent before stronger formula safety claims?
