# Release Handoff Blocking Actions

## Question

Can each top release implementation handoff be self-contained by carrying the owner actions that block that claim, instead of forcing owner loops to cross-reference global `nextOwnerActions`?

## Hypothesis

Yes. The release proof index already computes ranked owner actions. Embedding cloned, per-artifact `blockingActions` in each implementation handoff should make the top two owner loops actionable while keeping the global release gate fail-closed.

## External sources checked

- SLSA verifying artifacts: https://slsa.dev/spec/v1.0-rc1/verifying-artifacts
- SLSA v1.2 specification index: https://slsa.dev/spec/v1.2/
- GitHub CLI `gh attestation verify`: https://cli.github.com/manual/gh_attestation_verify
- GitHub artifact attestations: https://docs.github.com/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
- Sigstore bundle format: https://docs.sigstore.dev/about/bundle/
- Sigstore signature verification: https://docs.sigstore.dev/cosign/verifying/verify/

## Why this matters to Ascend

The top two claims are now release-proof handoffs, not research prompts. Owner loops need each handoff to state its own blockers, acceptance evidence, and forbidden shortcuts. Otherwise the safe-open and package-action owners can miss release constraints that are only visible in a global table. External provenance systems also reinforce that local digests must not be described as attestation unless verification material and identity/predicate expectations exist.

## Probe/implementation

Implemented a scoped proof-index fold-in:

- Added `blockingActions` to `ReleaseProofImplementationHandoff`.
- Populated it with the ranked owner actions for the handoff artifact.
- Cloned action objects when attaching them to handoffs so the global action list and per-handoff action list do not share mutable object references.
- Extended release-proof-index tests to assert:
  - safe-open blocking actions are attached in owner-priority order;
  - package-action blocking actions are attached in owner-priority order;
  - the compact owner handoff JSON includes per-claim blocking actions;
  - the release gate still blocks headline claims and surface promotion.

## Results

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
```

The targeted release-index tests pass with 4 tests and 140 assertions. The JSON handoff now gives safe-open and package-action owners their own `blockingActions` arrays, including product fixture acceptance, performance timing/streaming requirements, correctness boundary approval, and release provenance/publication policy.

## Confidence

High that this improves owner-loop actionability without changing product behavior. Medium that owners will keep the exact order, because the order reflects release-action priority rather than raw `readyWhen` declaration order.

## Fold-in decision

Promote to proof harness and owner-handoff data. This does not add SDK, CLI, API, MCP, or formula rename surfaces. The top two handoffs remain safe unknown workbook opening and auditable package-part mutation.

## Next question

Can the next owner loop resolve one product fixture gate by explicitly accepting generated structural fixtures, or should research keep searching for public binary replacements before release copy is considered?
