# Package Action Generated Fixture Acceptance Matrix

## Question

Can the package-action proof get the same generated-fixture acceptance matrix as safe-open, without weakening provenance or unsupported-feature boundaries?

## Hypothesis

Yes. The release proof index can make the generated signature-invalidation and unknown-part package cases owner-reviewable as local package-action evidence while still refusing provenance, signature preservation, arbitrary unknown-part preservation, and public-binary equivalence claims.

## External sources checked

- Microsoft digital signatures and code signing in workbooks: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing
- in-toto attestations: https://github.com/in-toto/attestation
- SheetJS write options and unsupported feature boundaries: https://docs.sheetjs.com/docs/api/write-options/
- openpyxl tutorial unsupported object warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html

## Why this matters to Ascend

Auditable package-part mutation is the rank-2 claim. The proof already demonstrates `passthrough`, `regenerate`, `add`, `drop`, and `error`, but the strongest edge cases still use generated package topology. The owner handoff needs to show exactly when those generated cases are acceptable as local package-action proof and when a public binary fixture is mandatory.

## Probe/implementation

Added `packageActionFixtureAcceptanceChecklist` to `releaseProofIndex.fixturePolicy` and rendered it in the Markdown release proof index. The checklist covers:

- `signature-invalidation-drop`: acceptable only as local evidence that generated edits drop or invalidate signature-related parts; public binary required for real signed workbook behavior, signature validity, re-signing, trust UX, or authoring provenance.
- `unknown-part-error`: acceptable only as local fail-closed package-action evidence with an explicit error action; public binary required for arbitrary third-party unknown-part preservation, understanding, recovery, or real workbook semantics.

Both rows have `gateEffect=keeps-edge-fixture-policy-missing-until-owner-approval`.

## Results

Focused validation passed:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bunx tsc --build
```

The owner handoff JSON now includes the package-action generated fixture acceptance checklist. It keeps `edge-fixture-policy` missing and does not change SDK, CLI, API, or MCP mutation surfaces.

## Confidence

High that the release proof index now separates generated package-action topology proof from public binary requirements. Medium that this closes enough ambiguity for product approval, because approval remains an owner decision.

## Fold-in decision

Promote to the release-proof index and product/release owner handoff. Do not promote provenance wording, unsupported-feature support, or a new package-action surface.

## Next question

Can the release claim board reduce the top two fixture-policy blockers into a single owner-ready acceptance packet with exact validation commands and forbidden wording?
