# Release Index Latency Exclusion

## Question

Should practical latency contract summaries be consumed by the release proof index now that they record worktree cleanliness?

## Hypothesis

Not yet. The release proof index should explicitly exclude latency contracts until they are generated from a tracked-clean worktree over standardized public inputs with product-approved threshold wording.

## External sources checked

- [SLSA source verification](https://slsa.dev/spec/v1.2/verifying-source) frames release verification around expectations for a specific source revision.
- [GitHub artifact attestations](https://docs.github.com/en/actions/concepts/security/artifact-attestations) describe signed provenance claims that bind build metadata such as repository and commit SHA to artifacts.
- [Hyperfine README](https://github.com/sharkdp/hyperfine) documents warmups and repeated benchmark runs, reinforcing that timing evidence needs explicit run context before comparison.

## Why this matters to Ascend

The release proof index is intentionally conservative. Adding local latency reports without a clean-source and public-input eligibility rule would weaken the claim board by making diagnostic measurements look like release evidence.

## Probe/implementation

Updated `fixtures/benchmarks/release-proof-index.ts` with a machine-readable `excludedEvidence` list. It now excludes `practical-latency-contracts` with:

- the reproduction command,
- the reason for exclusion,
- the eligibility rule,
- the owner loop,
- the honest boundary.

Updated `fixtures/benchmarks/release-proof-index.test.ts` to assert the exclusion metadata and Markdown rendering.

## Results

- `bun test fixtures/benchmarks/release-proof-index.test.ts` passed: 3 tests.
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings` rendered the new `Excluded Evidence` section with `practical-latency-contracts`.

## Confidence

High for the release-index guardrail. Medium for future latency publication readiness because it still needs a real clean-worktree, public-input benchmark run and product-approved threshold language.

## Fold-in decision

Promote to release/product loop as proof-index hygiene. Do not promote latency reports into release proof yet.

Allowed wording: "release proof packaging now records that practical latency contracts are excluded until clean public-input evidence exists."

## Next question

Can the top two release proof artifacts include enough fixture provenance metadata to distinguish generated edge packages from public binary fixtures without embedding large workbook data?
