# QSS Leapfrog Release Matrix

## Question

Can Ascend collapse its research/product portfolio into the top two QSS-leapfrog release claims, with every accepted evidence item linked to a committed proof artifact, test, benchmark, or RC gate?

## Hypothesis

A machine-readable release matrix inside the release proof index will be more useful than another prose synthesis because it can constrain owner loops, downgrade weak claims, and keep speculative research out of release positioning.

## External sources checked

- Quadratic getting started docs: https://docs.quadratichq.com/
- Quadratic spreadsheet navigation docs: https://docs.quadratichq.com/spreadsheet/navigating
- Quadratic formulas getting started docs: https://docs.quadratichq.com/formulas/getting-started
- Quadratic Python docs: https://docs.quadratichq.com/python/getting-started
- Quadratic SQL connection docs: https://docs.quadratichq.com/connections/sql-getting-started

## Why this matters to Ascend

Quadratic appears strongest as an interactive AI/code spreadsheet: Python, SQL, JavaScript, formulas, fast navigation, and database-connected analysis. Ascend should not counter-position with broad spreadsheet claims unless they are proven. The credible release posture is narrower: trust/proof/runtime for agentic workbook work, especially safe unknown workbook opening and auditable package-part mutation.

## Probe/implementation

Folded a `qssLeapfrogReleaseMatrix` into `fixtures/benchmarks/release-proof-index.ts` and the owner handoff JSON. The matrix is intentionally top-two-only:

- `safe-open-proof`: safe unknown workbook opening
- `package-action-proof`: auditable package-part mutation

Each row records what QSS likely does well, where Ascend is better only where evidence exists, accepted evidence, missing evidence, owner actions, forbidden claims, honest boundaries, and weak-claim dispositions. Active release blockers come from the existing readiness board. Deferred surfaces are archived research notes instead of release claims.

Validation added in `fixtures/benchmarks/release-proof-index.test.ts` checks that:

- the matrix has exactly two rows
- accepted evidence points to committed commands and paths
- weak claims become downgrade, blocker, or kill decisions
- active release blockers stay separate from archived research notes
- the owner handoff JSON carries the same matrix

## Results

The proof index now answers the release question directly:

- Allowed today: generated/local proof wording for safe-open and package-action only.
- Still blocked: public edge fixtures, release latency, package edge-fixture policy, unsupported-feature boundary approval, streaming matrix boundary, publication policy, and provenance boundary.
- Archived rather than promoted: formula language-service primitives, token-bounded agent view, retained viewport patch history, columnar scan sidecars, formula oracle routing, workflow observability, and practical latency contracts.

Commands run:

- `bunx tsc --build`
- `bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`
- `bun run test:changed` (expanded to the full suite; 5,210 pass, 1 skip, 1 timeout in `fixtures/benchmarks/competitive-real-workbook.test.ts`)
- `bun test fixtures/benchmarks/competitive-real-workbook.test.ts --timeout 120000` (64 pass, confirming the timed-out file passes when isolated with a longer timeout)

## Confidence

High for the release-claim shape and owner-routing behavior. Medium for QSS comparison wording because it relies on public Quadratic docs, not private product testing. Low for any claim that would require real-world public binary fixtures, signed provenance, broad latency thresholds, or direct competitive benchmark results.

## Fold-in decision

Fold into release, product, correctness, and performance owner loops as a release-claim gate. Do not promote new SDK/CLI/API/MCP surfaces from this work. The next implementation loops should close the top blockers or downgrade the corresponding claim wording.

## Next question

Which top blocker should an owner loop close first: product approval for disclosed generated topology fixtures, performance approval for safe-open release latency, or correctness approval for unsupported-feature package-action boundaries?
