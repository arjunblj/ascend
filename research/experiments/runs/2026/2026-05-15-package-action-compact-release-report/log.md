# Package Action Compact Release Report

Date: 2026-05-15

## Question

Can the package-action proof bundle expose a compact release report that includes the new readiness gate without embedding private workbook bytes or generated artifacts?

## Hypothesis

A compact report can preserve the product-shaped claim, readiness gates, case-level outcomes, and owner handoff facts while omitting the full proof payload. This should make the auditable package-part mutation claim easier to hand off without accidentally publishing raw workbook bytes, generated packages, or overbroad provenance language.

## External sources checked

- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- SLSA provenance model: https://slsa.dev/spec/v1.0-rc1/provenance
- in-toto attestation framework: https://github.com/in-toto/attestation
- openpyxl tutorial preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html

## Why this matters to Ascend

Ascend's North Star depends on trustworthy mutation planning. The full package-action proof is useful for engineers, but release/product owners need a smaller artifact that states the claim, evidence, and missing gates without becoming a fake attestation or leaking workbook payloads.

## Probe/implementation

I inspected `fixtures/benchmarks/package-action-proof.ts` and added `packageActionCompactReleaseReport(result)` plus a benchmark-only CLI mode:

```bash
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
```

The compact report includes:

- claim wording and explicit `headlineClaimAllowed=false`;
- release gate `blocked-by-publication-policy`;
- owner-loop `readyWhen` gates including `streaming-matrix-boundary`;
- source case counts and combined action counts;
- case-level action counts, journal issue counts, proof issue counts, and streaming summary counts;
- a boundary that excludes workbook bytes, per-part action rows, generated artifacts, signed provenance, SLSA, in-toto, third-party attestation, and full streaming parity.

It omits full proof rows, workbook bytes, output bytes, input digests, proof JSON byte counts, and streaming regenerated part paths.

## Results

The compact report preserves the core evidence shape:

- cases: 8
- source cases: 2 public fixtures, 2 generated workbooks, 4 generated edge packages
- action classes: passthrough, regenerate, add, drop, error
- post-write/proof issue case: `unknown-part-error`
- streaming proof cases: 1
- streaming regenerate part count: 1
- readyWhen gates: `edge-fixture-policy`, `provenance-boundary`, `unsupported-feature-boundary`, `streaming-matrix-boundary`

The test asserts the compact JSON is smaller than full proof JSON and does not contain `inputSha256`, `outputBytes`, `proofJsonBytes`, or `streamingRegeneratePartPaths`.

## Confidence

High for the artifact shape. Medium for release use because product/release owners still need to approve where this report lives and how generated edge packages are disclosed.

## Fold-in decision

Folded into the benchmark proof harness only. This is not a production mutation surface.

## Next question

Should the release proof index link to compact per-claim reports by digest, or should compact reports stay generated on demand until artifact storage and privacy policy are decided?
