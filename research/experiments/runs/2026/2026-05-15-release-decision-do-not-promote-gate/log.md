# Release Decision Do-Not-Promote Gate

## Question

Can the compact release-decision artifact carry the top two release claims and an explicit "do not promote yet" list, so owner loops do not accidentally turn archived research directions into release wording or implementation surfaces?

## Hypothesis

Yes. The existing QSS leapfrog matrix already separates top-two release rows from archived research notes. Folding the archived names into the compact release-decision JSON should make the claim gate stronger without adding a product surface.

## External sources checked

- Quadratic Docs: https://docs.quadratichq.com/
- Quadratic homepage: https://www.quadratichq.com/
- Microsoft calculation chain documentation: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain
- openpyxl tutorial: https://openpyxl.readthedocs.io/en/3.1/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

Ascend's current strongest release positioning is proof/runtime trust around safe opening and package mutation. Quadratic appears strong on AI, code cells, live data connections, and spreadsheet UX, so Ascend should avoid scattering release copy across speculative research directions until those directions change the top-two claim gates. A compact "do not promote yet" list gives implementation loops a machine-readable stop sign.

## Probe/implementation

- Added `doNotPromoteYet` to `ReleaseProofReleaseDecisionBoard`.
- Populated it from `qssLeapfrogReleaseMatrix.archivedResearchNotes`.
- Kept `--release-decision-json` compact: it still omits `claimBlockerBoard`, `fixturePolicy`, and `deferredClaims`.
- Rendered the same list in the Markdown release proof index under the Release Decision Board.
- Added tests that pin the archived names, `do-not-promote-yet` status, compact JSON shape, and Markdown section.

## Results

- `bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`: 5 pass, 0 fail.
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --release-decision-json`: passed and emitted top-two rows plus seven do-not-promote entries.
- `bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts research/experiments/index.md research/experiments/runs/2026/2026-05-15-release-decision-do-not-promote-gate/log.md`: passed.
- `bunx tsc --build`: passed.
- `bun run test:changed`: 5219 pass, 1 skip, 0 fail.

The gate now says:

- safe-open and package-action remain the only release-decision rows.
- formula language-service primitives, token-bounded agent view, retained viewport patch history, columnar scan sidecars, formula oracle routing, agent workflow observability, and practical latency contracts remain do-not-promote until they change the top-two claim gate.

## Confidence

High for gate shape and claim stewardship. Medium for the competitive boundary because it relies on public QSS documentation rather than a hands-on QSS probe, which is acceptable here because the change is a release claim throttle rather than a behavior comparison.

## Fold-in decision

Promote to release loop. This is a small benchmark/release-proof harness fold-in, not a production workbook behavior change.

## Next question

Can the product owner fixture gate be made similarly compact by producing a one-command acceptance packet for generated signed/unknown topology fixtures versus public binary replacements?
