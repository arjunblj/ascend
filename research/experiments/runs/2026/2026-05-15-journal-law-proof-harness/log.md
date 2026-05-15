# Journal Law Proof Harness

Date: 2026-05-15

## Question

Can property-style journal-law proof move from an ignored research probe toward a tracked correctness artifact without adding a new dependency yet?

## Hypothesis

Yes. A deterministic generated harness can prove a useful subset now: exact inverse restoration across multiple public operation families, plus explicit lossy classification for data-validation and conditional-format metadata order/duplicate boundaries. It should not claim full property-based testing until shrinkable generation is available.

## External sources checked

- fast-check model-based testing defines command sequences with `check`, `run`, and shrinkable replay paths, which is the right future shape for workbook operation sequences: https://fast-check.dev/docs/advanced/model-based-testing/
- fast-check documents Bun test runner setup, confirming a future dependency can fit Ascend's current test stack: https://fast-check.dev/docs/tutorials/setting-up-your-test-environment/
- Hypothesis stateful testing generates sequences of primitive actions and searches for failures, which matches journal-law testing better than isolated examples: https://hypothesis.readthedocs.io/en/latest/stateful.html
- QuickCheck frames properties as executable claims checked over generated inputs; Ascend's immediate property is "exact journals restore workbook evidence after inverse apply": https://www.cse.chalmers.se/~rjmh/QuickCheck/manual.html

## Why this matters to Ascend

The auditable mutation claim is only credible if inverse journals are tested as laws, not only as handpicked examples. Recent metadata-order and metadata-duplicate fixes show the risk: public inverse operations can restore content while failing to restore order or selector identity. A generated harness keeps those boundaries visible.

## Probe/implementation

Folded in a tracked proof harness:

- Added `fixtures/benchmarks/journal-law-proof.ts`.
- Added `fixtures/benchmarks/journal-law-proof.test.ts`.
- The harness generates deterministic exact operation sequences over cells, formulas, comments, hyperlinks, freeze panes, data validations, and conditional formats.
- The harness also seeds existing row layout, column layout, sheet protection, tab color, page setup, and print-area metadata before proving their replacement inverses are exact.
- For exact generated sequences, it applies operations with `journal: true`, applies `journal.inverseOps` transactionally, and compares stable workbook evidence before and after.
- It separately probes lossy metadata boundaries:
  - non-suffix data-validation delete;
  - duplicate data-validation delete;
  - non-tail conditional-format delete;
  - non-tail conditional-format replacement;
  - duplicate conditional-format delete.

The first generator draft was too broad: it included known lossy operation families such as created row/column layout, absent workbook protection, page setup, tab color, and package-state formats. That failure was useful evidence, so the tracked harness now distinguishes exact-law generation from explicit lossy-boundary cases.

## Results

Latest proof:

```bash
bun run fixtures/benchmarks/journal-law-proof.ts
```

Observed:

- 58 total cases.
- 53 exact law cases: 48 generated exact sequences plus 5 pre-seeded metadata replacement cases.
- 5 lossy boundary cases.
- 0 failures.
- Exact operation counts included:
  - `setCells=35`
  - `setFormula=34`
  - `setComment=34`
  - `setHyperlink=34`
  - `freezePane=34`
  - `setDataValidation=34`
  - `setConditionalFormat=36`
  - `setRowHeight=1`
  - `hideRows=1`
  - `setColWidth=1`
  - `hideCols=1`
  - `setSheetProtection=1`
  - `setTabColor=1`
  - `setPageSetup=1`
  - `setPrintArea=1`
- Lossy issue counts included:
  - `data-validations:metadata-order=1`
  - `data-validations:metadata-duplicate=1`
  - `conditional-formats:metadata-order=2`
  - `conditional-formats:metadata-duplicate=1`

Validation passed:

```bash
bun run fixtures/benchmarks/journal-law-proof.ts --json
bun test fixtures/benchmarks/journal-law-proof.test.ts
bunx biome check --write fixtures/benchmarks/journal-law-proof.ts fixtures/benchmarks/journal-law-proof.test.ts
bunx tsc --build
```

## Confidence

Medium-high for the covered exact operation families, the pre-seeded row/column/page/protection/tab metadata replacement families, and metadata lossy boundaries. Medium overall because the generator is deterministic and not shrinkable. It proves a stronger subset than hand fixtures, but it is not yet a full fast-check model-based test.

## Fold-in decision

Promote to correctness proof harness and commit. Keep it out of release proof index until it uses broader operation generators and shrinkable counterexamples.

## Next question

Should the next correctness loop add `fast-check` and convert this deterministic proof into a shrinkable model-based `bun test`, or should it first broaden exact families into package-state and style-related operations without collapsing known creation-loss boundaries into exact cases?
