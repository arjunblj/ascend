# Safe Open Generated Fixture Acceptance Matrix

## Question

Can the safe-open proof owner handoff distinguish topology-only generated package proof from real-world public binary fixture requirements in machine-readable form?

## Hypothesis

Yes. The release proof index should carry a case-level acceptance matrix for the generated safe-open cases. The matrix should explain when `signed`, `unknown-part`, and `malformed` generated cases are acceptable as topology proof, when public binaries are required, and that the product gate remains missing until owner approval.

## External sources checked

- GitHub repository license API: https://docs.github.com/en/rest/licenses/licenses
- GitHub large-file guidance: https://docs.github.com/github/managing-large-files/working-with-large-files
- OpenSSF Scorecard binary artifacts check: https://github.com/ossf/scorecard/blob/main/docs/checks.md
- SLSA provenance specification: https://slsa.dev/spec/v1.0-rc1/provenance

## Why this matters to Ascend

Safe unknown workbook opening can be a credible product claim only if generated edge packages are disclosed honestly. Owners need a checklist that prevents generated package topology from being mistaken for public binary fixture coverage, signature validity, vendor behavior, or file-safety evidence.

## Probe/implementation

Added `safeOpenFixtureAcceptanceChecklist` to `releaseProofIndex.fixturePolicy` and rendered it in the Markdown release proof index. The checklist covers:

- `signed`: acceptable only as signature-related package-topology proof; public binary required for real signed workbook behavior, signature validity, vendor repair UX, or trust wording.
- `unknown-part`: acceptable only as unknown-feature review-routing topology proof; public binary required for arbitrary third-party unknown-part preservation or real workbook semantics.
- `malformed`: acceptable only as fail-closed rejection-path proof; public binary required for vendor repair equivalence or arbitrary malformed workbook recovery claims.

Every row has `gateEffect=keeps-public-edge-fixtures-missing-until-owner-approval`.

## Results

Focused validation passed:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bunx tsc --build
```

The owner handoff JSON now includes the three-row safe-open generated fixture acceptance checklist. This is proof packaging only; it does not mark `public-edge-fixtures` satisfied.

## Confidence

High that the release proof index now separates generated topology proof from public binary fixture requirements. Medium that this is sufficient for product approval, because the approval itself remains an owner decision.

## Fold-in decision

Promote to the release-proof index and product-owner handoff. Do not promote a new open surface, do not close the safe-open product gate, and do not change claim wording beyond clearer owner evidence.

## Next question

Can the package-action proof get the same generated-fixture acceptance matrix for signature invalidation and unknown-part error without weakening provenance or unsupported-feature boundaries?
