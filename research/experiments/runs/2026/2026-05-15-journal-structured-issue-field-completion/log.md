# Journal Structured Issue Field Completion

## Question

Can the remaining in-flight journal issue edit be safely validated so all emitted lossy/unsupported issues carry v1 `surface` and `reason` fields?

## Hypothesis

Yes. The diff is a tiny correctness fix that adds missing structured fields to existing issue objects without changing operation semantics.

## External sources checked

None in this cycle. This was pure local validation of an already in-flight SDK schema fix.

## Why this matters to Ascend

The auditable mutation claim depends on machine-readable journal issue reasons. If some SDK paths still emit issue objects without `surface` and `reason`, the v1 schema is weaker than the release proof language and clients must fall back to inference.

## Probe/implementation

Inspected the only remaining dirty production file, `packages/sdk/src/journal.ts`. The in-flight diff, later landed as `7ab27f71 fix(sdk): structure remaining journal preimage issues`, adds explicit `surface` and `reason` fields to:

- deleted-sheet restore failure;
- conditional-format delete order loss;
- data-validation move order loss;
- table column restore failure;
- deleted table restore failure;
- missing table sheet restore failure.

No operation behavior changed.

Validation:

```bash
bun test packages/sdk/src/journal-exactness.test.ts packages/sdk/src/journal-compatibility.test.ts packages/sdk/src/workbook-property-journal.test.ts
bunx biome check packages/sdk/src/journal.ts
bunx tsc --build
```

## Results

Focused journal validation passed:

- 23 journal exactness/compatibility tests passed.
- Biome check passed.
- Typecheck passed.

The public issue schema already requires `code`, `message`, `surface`, and `reason`; the fix completes direct emission paths that otherwise relied on classifier inference.

## Confidence

High. This is a metadata-only structured issue completion with focused tests over the journal taxonomy and compatibility schema.

## Fold-in decision

Promote to correctness loop. The tiny SDK fix is now committed as `7ab27f71`; keep this log as validation bookkeeping and return to claim stewardship. Do not broaden journal semantics in this block.

## Next question

Can the top-two proof owners now receive concise next-loop prompts without adding more surfaces?
