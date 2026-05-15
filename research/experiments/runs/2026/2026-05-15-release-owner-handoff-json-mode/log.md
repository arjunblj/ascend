# Release Owner Handoff JSON Mode

## Question

Can owner loops consume top release handoffs, deferred claims, and excluded evidence without parsing the full release proof artifact index?

## Hypothesis

Yes. A compact `--owner-handoffs-json` mode can expose only the routing data owners need: release gate status, promotion gate, implementation handoffs, deferred claims, excluded evidence, and boundary text.

## External sources checked

- SLSA provenance uses structured predicates and subjects for downstream verification, reinforcing that owner loops should consume structured fields rather than scrape prose: https://slsa.dev/spec/v1.0-rc1/provenance
- in-toto attestation metadata is structured around subject and predicate fields, reinforcing a compact machine-readable handoff shape while keeping Ascend below attestation claims: https://github.com/in-toto/attestation/blob/main/spec/README.md
- OpenSSF Scorecard supports machine-readable JSON output for automated consumers, supporting a separate compact output mode for proof routing: https://github.com/ossf/scorecard

## Why this matters to Ascend

The claim steward loop has narrowed implementation handoff to two product-shaped claims. Owner loops should not need to parse full proof artifacts or Markdown synthesis to learn what to do next. A compact JSON mode reduces drift and supports future automation while preserving the boundary that this is not a product API or signed release bundle.

## Probe/implementation

Added `ReleaseProofOwnerHandoffIndex` and `releaseProofOwnerHandoffIndex()` to `fixtures/benchmarks/release-proof-index.ts`.

Added CLI mode:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
```

The compact output includes:

- release gate status;
- headline and implementation surface promotion flags;
- missing requirement count;
- top-two `implementationHandoffs`;
- `deferredClaims`;
- `excludedEvidence`;
- boundary text.

The focused test asserts that this output includes top handoffs and deferred/excluded evidence while not embedding the full `artifacts` array.

## Results

- Focused test passed: 4 tests, 88 assertions.
- CLI spot check showed `implementationHandoffs`, `deferredClaims`, and `excludedEvidence`.
- The compact output excludes the full proof artifact payloads.
- `implementationSurfacePromotionAllowed` remains false.

## Confidence

High for output shape and routing value. Medium for future owner-loop consumption until those loops are updated to call `--owner-handoffs-json` directly.

## Fold-in decision

Fold into the release proof harness. This is a benchmark/proof routing surface only, not SDK/CLI/API/MCP product surface.

## Next question

Can product and release owners resolve the generated-fixture acceptance gate directly from the compact handoff output?
