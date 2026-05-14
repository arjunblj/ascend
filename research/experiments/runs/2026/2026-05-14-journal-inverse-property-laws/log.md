# Journal Inverse Property Laws

## Question

Can property-based operation testing prove inverse journal laws better than hand fixtures?

## Hypothesis

Yes. The journal already has broad example coverage, but generated operation sequences can expose interactions humans are unlikely to author by hand. The core law is: if a journal claims `exact: true`, then `apply(ops); apply(journal.inverseOps)` must restore a comparable workbook state. If the law cannot hold, the journal must classify the sequence as lossy or unsupported.

## External sources checked

- [fast-check](https://fast-check.dev/): property-based testing for JavaScript and TypeScript.
- [fast-check model-based testing](https://fast-check.dev/docs/advanced/model-based-testing/): command generators with `check` and `run` map well to workbook operation sequences.
- [fast-check Bun setup](https://fast-check.dev/docs/tutorials/setting-up-your-test-environment/): confirms a future production harness can fit the current Bun test stack.
- [Hypothesis stateful testing](https://hypothesis.readthedocs.io/en/latest/stateful.html): state machines generate whole test programs, not just individual inputs.
- [QuickCheck paper](https://www.researchgate.net/publication/2449938_QuickCheck_A_Lightweight_Tool_for_Random_Testing_of_Haskell_Programs): property-based testing asks developers to state properties and lets random generation search for counterexamples.

## Why this matters to Ascend

Preview, commit, undo, repair, and agent approval all rely on journal honesty. For agent-native editing, `exact` must be a claim with evidence, not a convenience flag. A generated inverse-law harness can turn journal correctness into a replayable proof artifact and produce shrinkable prompts for the correctness loop.

## Probe/implementation

Inspected local implementation:

- `packages/sdk/src/journal.ts` builds `MutationJournal` entries, inverse ops, exact/support flags, and lossy issues.
- `packages/sdk/src/workbook.ts` applies operations with `journal: true`.
- `packages/sdk/src/interactive-contract.test.ts` has extensive example-based journal coverage across cells, formulas, styles, comments, tables, structural edits, workbook metadata, and lossy metadata classification.

Added ignored probe `research/experiments/runs/2026/2026-05-14-journal-inverse-property-laws/probes/generated-journal-laws.ts`. It uses a deterministic PRNG, creates 120 eight-operation sequences over a seeded workbook, applies each sequence with `journal: true`, applies inverse ops for exact journals, and compares workbook state with volatile snapshot timestamps removed.

The generated command set includes:

- `setCells`
- `setFormula`
- `clearRange`
- `setNumberFormat`
- `setStyle`
- `insertRows` / `deleteRows`
- `insertCols` / `deleteCols`

## Results

The first probe version found 66 failures, but the first failure was a probe bug: `wb.snapshot()` includes a timestamp. After removing volatile timestamps, 14 exact-journal failures remained.

The minimal counterexample was a single operation:

```ts
{ op: 'setStyle', sheet: 'Sheet1', range: 'E9', style: { font: { bold: true }, numberFormat: '0.0' } }
```

Before the fix, the journal reported `supported: true`, `exact: true`, and `inverseOps: []`, so undo left a style-only empty cell at `E9`.

Folded in a small production fix:

- `packages/sdk/src/journal.ts`: `styleInverseOps` now emits `clearRange(..., what: 'all')` for non-existent cell preimages, removing style-only cells created by style operations.
- `packages/sdk/src/interactive-contract.test.ts`: added `journal inverse ops clear styles created on empty cells`.

After the fix, the generated probe reported:

| Metric | Count |
| --- | ---: |
| Generated cases | 120 |
| Exact cases checked | 86 |
| Classified lossy/unsupported cases | 34 |
| Exact inverse failures | 0 |
| Exact inverse ops applied | 1716 |

Validation:

- `bun run research/experiments/runs/2026/2026-05-14-journal-inverse-property-laws/probes/generated-journal-laws.ts` passed with 0 failures.
- `bun test packages/sdk/src/interactive-contract.test.ts -t "journal inverse ops clear styles created on empty cells"` passed.
- `bun test packages/sdk/src/interactive-contract.test.ts -t "journal"` passed: 63 tests, 0 failures.
- `bunx biome check packages/sdk/src/journal.ts packages/sdk/src/interactive-contract.test.ts` passed.
- `bunx tsc --build` passed.
- `bun run test:changed` ran the full suite because of repository-level changes: 4850 pass, 1 skip, 0 fail.

## Confidence

High that property-style generated sequences are valuable for Ascend's correctness loop. This cycle found and fixed a real inverse-law bug with a small generated harness and a minimal focused test.

## Fold-in decision

Folded in the minimal correctness fix now.

Promote the generated probe to the correctness loop as a production `bun test` harness, ideally using `fast-check` for shrinking once dependency policy is decided.

Do not promote to performance or product/DX yet. The immediate value is correctness evidence and counterexample generation.

## Next question

Can real workbook oracle routing choose Excel, LibreOffice, HyperFormula, static goldens, or accepted mismatch classes automatically?
