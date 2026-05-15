# Release Decision Board Gate

## Question

Can the top-two release-claim decision artifact be made machine-readable without opening new research or product surfaces?

## Hypothesis

The release-proof index already has enough evidence to derive a compact decision board for only the top two claims. Adding that board to local proof output should reduce overclaiming risk by pinning allowed wording, exact proof links, forbidden claims, and owner blockers in tests.

## External sources checked

- Quadratic docs: current product positioning around AI, Python, SQL, JavaScript, formulas, and database connections. <https://docs.quadratichq.com/>
- Quadratic navigating docs: current claim around browser spreadsheet navigation and 60 FPS WASM/WebGL interaction. <https://docs.quadratichq.com/spreadsheet/navigating>
- SLSA 1.2 build provenance distribution: provenance is distributed as attestations bound to artifacts, not implied by local proof reports. <https://slsa.dev/spec/v1.2/distributing-provenance>
- GitHub artifact attestations: build provenance requires explicit artifact attestation workflows, not local digest output. <https://docs.github.com/en/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds>

## Why this matters to Ascend

Ascend's release message should be the trust/proof/runtime layer for agentic spreadsheet work, but only where current evidence supports it. A machine-readable board lets implementation, product, performance, correctness, and release loops consume the same claim boundaries instead of relying on a prose handoff.

## Probe/implementation

- Added `releaseDecisionBoard` to `fixtures/benchmarks/release-proof-index.ts`.
- The board has exactly two rows: `safe-open-proof` and `package-action-proof`.
- Each row carries:
  - `claimWordingAllowedToday`
  - `headlineClaimAllowed`
  - `implementationSurfacePromotionAllowed`
  - `proofRequired`
  - `acceptedEvidence`
  - `claimsWeMustNotMake`
  - `aPlusBlockingOwnerActions`
  - boundary text
- Included the board in owner-handoff JSON and Markdown output.
- Added regression tests proving the board does not allow headline claims or implementation promotion while release gates remain blocked.
- Added `--release-decision-json` so owner loops can fetch only the compact two-claim board without the full owner-handoff payload.
- Updated release provenance references to current SLSA 1.2 distribution guidance and test-pinned the fixture policy away from old SLSA v1.0 URLs.

## Results

- `bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`: passed, 5 tests.
- `bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`: passed.
- `bunx tsc --build`: passed.
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`: passed and emits the owner handoff without embedding full proof artifacts.
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --release-decision-json`: passed and emits only the compact release decision board.
- `bun run test:changed`: passed, 5218 pass, 1 skip, 0 fail.

## Confidence

High. This is a derived local proof artifact with test-pinned boundaries; it does not add product behavior or loosen any gate.

## Fold-in decision

Promote to release/product/correctness/performance owner routing only. Do not promote new SDK, CLI, API, MCP, benchmark-threshold, or public claim surfaces from this artifact.

## Next question

Stay collapsed. The next implementation loop should resolve one existing top-two blocker: safe-open fixture/latency owner approval, package-action generated-edge fixture policy, unsupported-feature boundary approval, streaming matrix owner approval, or compact-report publication policy.
