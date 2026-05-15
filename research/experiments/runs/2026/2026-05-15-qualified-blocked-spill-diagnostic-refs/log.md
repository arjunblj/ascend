# Qualified Blocked-Spill Diagnostic Refs

## Question

Can an in-flight copy-sheet dynamic/spill regression test be completed without expanding formula-intelligence scope or promoting rename?

## Hypothesis

Yes. The useful correction is narrowly in verifier diagnostics: blocked-spill diagnostics should display qualified binding refs without duplicating sheet names. The copy-sheet quality moat should assert retargeted dynamic/spill metadata and clean post-write audits, not turn blocked spill warnings into a new product surface.

## External sources checked

None for this cycle. This was a local validation/fix cycle over an already dirty in-flight regression test.

## Why this matters to Ascend

Auditable mutation and formula trust both depend on refs being exact and readable. A diagnostic like `Copy!Copy!A5` makes agent repair evidence less trustworthy even when the underlying workbook metadata is structurally valid.

## Probe/implementation

Ran the in-flight focused test:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "quality moat matrix"
```

It failed while completing the copied dynamic/spill case. A local probe also showed that qualified blocked-spill refs could render with a duplicated sheet name in verifier diagnostics.

Implemented a scoped verifier formatting fix in `packages/verify/src/checker.ts`:

- format blocked-spill ranges through parsed binding ranges;
- format blocking refs through parsed binding cell refs;
- preserve fallback behavior for unparseable refs.

Completed the in-flight SDK quality moat test and added a focused verifier regression for qualified blocked-spill diagnostic refs.

## Results

Focused validation:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "quality moat matrix"
bun test packages/verify/src/verify.test.ts -t "blocked spill"
bunx biome check packages/verify/src/checker.ts packages/verify/src/verify.test.ts packages/sdk/src/agent-workflow.test.ts
bunx tsc --build
bun run test:changed
```

Results:

- SDK quality moat: 1 pass, 0 fail.
- Verify blocked-spill subset: 7 pass, 0 fail.
- Biome passed after formatting.
- Typecheck passed.
- `test:changed` passed: 4393 pass, 1 skip, 0 fail, 26751 expect calls.
- Qualified blocked-spill diagnostics now render `Copy!A4:A5` and `Copy!A5`, not duplicated sheet-qualified refs.

## Confidence

High for the diagnostic formatting fix and copied dynamic/spill quality moat. Medium for broader formula-binding persistence because this does not change writer semantics or claim safe rename.

## Fold-in decision

Promote to correctness loop as hygiene under formula trust and auditable mutation. Do not promote a new formula intelligence surface and do not implement rename.

## Next question

Return to claim stewardship: do the current top-two release handoffs need owner approval, fixture replacement, or publication policy before research should run another implementation probe?
