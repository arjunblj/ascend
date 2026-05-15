# Release Trust Completeness Boundary

## Question

Can the release trust matrix say what is deliberately out of scope without reopening broad research or implying that headline release claims are approved?

## Hypothesis

If the matrix pins a small out-of-scope table with tests, then correctness loops can reject unrelated expansion work unless it names a silent corruption, exact-journal, or post-write drift path tied to the top release claims. The wording must stay below product, performance, fixture-policy, provenance, and publication gates.

## External sources checked

- SLSA provenance and build-claim boundary: https://slsa.dev/spec/v1.0-rc1/provenance
- SLSA provenance distribution guidance: https://slsa.dev/spec/v1.0/distributing-provenance
- GitHub artifact attestations for build provenance: https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
- Microsoft Protected View boundary for untrusted Office files: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653

## Why this matters to Ascend

The top two release claims are already blocked by owner gates. Correctness work should not expand into generic formula breadth, UX orchestration, performance tuning, malformed-field enumeration, or unknown-feature implementation unless it changes a concrete trust failure path. This keeps the QSS-leapfrog matrix collapsed around proof and prevents overclaiming.

## Probe/implementation

Finished the in-flight `docs/RELEASE_TRUST_MATRIX.md` boundary edit and tightened one phrase from broad "release claim" wording to "correctness/trust slice of release wording." Added a pinned test in `packages/sdk/src/release-trust-matrix.test.ts` that parses the `Completeness Boundary` table and freezes the five out-of-scope classes.

No SDK, CLI, API, MCP, formula, viewport, storage, or benchmark surface was added.

## Results

The matrix now explicitly says it does not close product, performance, fixture-policy, provenance, or publication gates. It also freezes these out-of-scope classes unless a concrete bug proves release trust impact:

- broad formula function coverage
- product/DX orchestration such as progressive open or viewport merge helpers
- reader/writer performance and benchmark tuning
- more malformed-field enumeration
- new unknown Excel feature implementation

Validation:

- `bun test packages/sdk/src/release-trust-matrix.test.ts --timeout 30000`
- `bunx biome check packages/sdk/src/release-trust-matrix.test.ts`
- `bunx tsc --build`
- `git diff --check -- docs/RELEASE_TRUST_MATRIX.md packages/sdk/src/release-trust-matrix.test.ts`

## Confidence

High that the change narrows correctness scope and prevents overclaiming. Medium that the exact out-of-scope list is complete; it is intentionally minimal and test-pinned for this release block.

## Fold-in decision

Fold into release documentation/test hygiene only. This changes claim-boundary behavior for owner loops but does not satisfy headline release gates.

## Next question

Should the release proof index expose this correctness/trust completeness boundary in owner handoff JSON, or is the pinned release trust matrix enough for correctness owners?
