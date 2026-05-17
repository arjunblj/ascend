# Fixture Decision Packet

## Question

Can the product fixture gate be collapsed into a one-command packet that tells the owner exactly which generated edge fixtures are disclosed, which public replacements are still missing, and what shortcuts remain forbidden?

## Hypothesis

Yes. The release proof index already contains fixture policy, tracked scan evidence, and generated fixture decision rows. A compact `--fixture-decision-json` mode can expose only the product decision packet, without leaking the broader owner handoff or implying that the gate is approved.

## External sources checked

- GitHub repository limits: https://docs.github.com/en/repositories/creating-and-managing-repositories/repository-limits
- GitHub large files: https://docs.github.com/github/managing-large-files/working-with-large-files
- OpenSSF Scorecard binary-artifacts guidance: https://github.com/ossf/scorecard/blob/main/docs/checks.md
- SLSA 1.2 provenance distribution: https://slsa.dev/spec/v1.2/distributing-provenance

## Why this matters to Ascend

The top two release claims are blocked by product approval of generated signed/unknown topology fixtures or replacement with public binary fixtures. The owner should not have to inspect the full release proof index to decide this. A compact packet keeps the release matrix collapsed around safe-open and package-action proof while making the next owner action reviewable.

## Probe/implementation

- Added `ReleaseProofFixtureDecisionPacket`.
- Added `releaseProofFixtureDecisionPacket(result)`.
- Added `--fixture-decision-json` to `fixtures/benchmarks/release-proof-index.ts`.
- Included only product-owned fixture approval rows, tracked scan summaries, generated fixture cases, validation commands, source references, and forbidden shortcuts.
- Added a regression test that verifies the compact output omits broader `claimBlockerBoard`, `deferredClaims`, and `qssLeapfrogReleaseMatrix` payloads.

## Results

- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --fixture-decision-json`: passed and emitted the product decision packet.
- `bun test fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`: 9 pass, 0 fail.
- `bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json`: 224 tracked fixtures scanned, public signed replacement still missing.
- `bun run fixtures/benchmarks/package-action-fixture-scan.ts --json`: 224 tracked fixtures scanned, `signaturePackage=0` and `unknownPathFamily=1`.

## Confidence

High for shape and owner-routing value. Medium for external fixture completeness because the scan is intentionally limited to the tracked public corpus and does not prove that no suitable public workbook exists elsewhere.

## Fold-in decision

Promote to release/product loop. This is a small release-proof harness fold-in that improves owner review; it does not approve generated fixtures, publish workbook binaries, or add production workbook behavior.

## Next question

Can the correctness unsupported-feature boundary be collapsed into the same style of one-command owner packet for package-action wording?
