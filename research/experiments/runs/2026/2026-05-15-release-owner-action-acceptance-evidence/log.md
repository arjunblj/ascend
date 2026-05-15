# Release Owner Action Acceptance Evidence

## Question

Can the top-claim owner handoff become actionable without adding product surfaces by making each missing release-readiness blocker state the acceptance evidence and forbidden shortcut in machine-readable proof output?

## Hypothesis

Yes. `release-proof-index` already ranks owner actions and blocks headline claims. Adding explicit `acceptanceEvidence` and `forbiddenShortcut` fields to each next owner action should reduce ambiguity for product, correctness, performance, and release loops while keeping `headlineClaimsAllowed=false` and `implementationSurfacePromotionAllowed=false`.

## External sources checked

- SLSA provenance: https://slsa.dev/spec/v0.2/provenance
- GitHub artifact attestations: https://docs.github.com/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds
- Sigstore bundle format: https://docs.sigstore.dev/about/bundle/
- Sigstore verification: https://docs.sigstore.dev/cosign/verifying/verify/

## Why this matters to Ascend

The claim board has correctly stopped new surface promotion, but owner loops still need a concise, machine-readable checklist for what would unblock the top two claims. Without explicit acceptance and forbidden-shortcut fields, a future loop could accidentally treat local digests, generated package topology, or one streaming proof as stronger release evidence than they are.

## Probe/implementation

Implemented a scoped benchmark/proof harness fold-in:

- Added `acceptanceEvidence` and `forbiddenShortcut` to `ReleaseProofNextOwnerAction`.
- Populated those fields for fixture policy, unsupported-feature boundary, release-latency validation, streaming matrix boundary, provenance/publication boundary, and compact-report publication policy.
- Included `nextOwnerActions` in `releaseProofOwnerHandoffIndex`, so `--owner-handoffs-json` carries the actionable criteria without embedding full proof artifacts.
- Updated release-proof-index tests for the new fields and compact owner handoff JSON.

## Results

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
```

The targeted test passed with 4 tests and 95 assertions. Biome passed. The owner-handoff JSON now includes `nextOwnerActions` with per-blocker `acceptanceEvidence` and `forbiddenShortcut`, for example:

- Product fixture acceptance must either approve disclosed generated structural packages or replace them with public binary fixtures.
- Performance latency acceptance must use tracked-clean release-environment evidence over standardized public inputs.
- Release provenance acceptance must keep local proof below SLSA, in-toto, Sigstore, GitHub artifact attestation, and signed-provenance thresholds.
- Compact report publication must wait for storage, retention/privacy filtering, canonicalization, and verification expectations.

## Confidence

High that this makes the existing top-two owner handoff more actionable without changing product behavior. Medium that the exact wording is final; owner loops may still refine the policy text, but the release gate remains fail-closed until they do.

## Fold-in decision

Promote to proof harness. This is not a production surface and does not authorize SDK, CLI, API, or MCP promotion. Keep formula rename frozen and keep lower-ranked claims out of the release proof index.

## Next question

Should the safe-open and package-action compact Markdown reports render the same acceptance evidence, or is keeping it in `release-proof-index --owner-handoffs-json` enough for owner-loop execution?
