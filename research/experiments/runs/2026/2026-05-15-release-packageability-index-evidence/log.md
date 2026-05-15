# Release Packageability Index Evidence

Date: 2026-05-15

## Question

Can the machine-readable release proof index expose the new SDK and app tarball smoke evidence without implying publication, signing, or full protocol readiness?

## Hypothesis

If packageability evidence is represented as a first-class release-index field with commands, covered evidence, missing policy requirements, forbidden claims, and a boundary, then owner loops can consume the proof without reading research prose or interpreting tarball smoke as provenance.

## External sources checked

- npm bundled dependencies documentation: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#bundleddependencies
- npm package `bin` documentation: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#bin
- SLSA provenance specification: https://slsa.dev/spec/v1.0/provenance
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

The release claim board is meant to be machine-readable. The app and SDK tarball smokes are now real proof, but they should not unlock headline claims until publication policy, provenance wording, API lifecycle smoke, and MCP stdio protocol smoke are owner-approved.

## Probe/implementation

- Added `releasePackageabilityEvidence` to `ReleaseProofIndexResult`.
- Added the same field to `ReleaseProofOwnerHandoffIndex`.
- Rendered the evidence in release proof Markdown.
- Added tests asserting commands, covered evidence, missing policy requirements, forbidden claims, and Markdown visibility.

## Results

Proof commands run:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bunx tsc --build
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json | jq '{releasePackageabilityEvidence, artifactCount, excludedEvidenceCount}'
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json | jq '{releasePackageabilityEvidence, releaseGate, headlineClaimsAllowed, implementationSurfacePromotionAllowed}'
```

Observed evidence:

- Release-index tests passed: 4 tests, 351 expectations.
- Full JSON and owner-handoff JSON now include `releasePackageabilityEvidence`.
- The field names `bun run release:sdk:smoke` and `bun run release:apps:smoke`.
- The gate remains fail-closed: `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, and `releaseGate=blocked-by-publication-policy`.

## Confidence

High for machine-readable routing. Medium for release readiness because this records local tarball smokes only, not registry publication or signed provenance.

## Fold-in decision

Promote to release/product loop as proof-index reporting. This is not a new product surface and does not change top claim handoffs.

## Next question

Should release ownership next define artifact storage/canonicalization policy or implement a stdio MCP protocol-session smoke from the packed `@ascend/mcp` tarball?
