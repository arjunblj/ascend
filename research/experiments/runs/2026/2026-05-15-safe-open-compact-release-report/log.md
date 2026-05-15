# Safe Open Compact Release Report

Date: 2026-05-15

## Question

Can the safe-open proof get the same compact release-report treatment without adding a new open surface or weakening the public-fixture blocker?

## Hypothesis

Yes. A compact safe-open report can expose the product claim, review routing, case mix, malformed rejection, owner-loop blockers, and honest boundaries without embedding workbook bytes or input digests.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft OPC digital signatures: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/digital-signatures
- SheetJS parse options: https://docs.sheetjs.com/docs/api/parse-options

## Why this matters to Ascend

Safe unknown workbook opening is the top product/performance claim. The full proof is useful for engineers, but release owners need a small artifact that states exactly what is proven and what still blocks stronger wording.

## Probe/implementation

I added `safeOpenCompactReleaseReport(result)` and a benchmark-only CLI mode:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json
```

The compact report includes:

- claim wording and `headlineClaimAllowed=false`;
- release gate `blocked-by-publication-policy`;
- owner-loop `readyWhen` blockers: `public-edge-fixtures`, `release-latency-run`, `publication-boundary`;
- case-kind counts, recommended mode counts, risk families, malformed rejection, and compact case rows;
- boundary language excluding malware scanning, sandboxing, file trust, active-content safety, signed provenance, and release performance thresholds.

It omits workbook bytes and input digests.

## Results

The focused proof test now verifies:

- 9 cases, 8 OK, 1 rejected;
- 6 public fixture cases, 2 synthetic cases, 1 malformed case;
- 4 review-before-hydration cases;
- 4 `formula` recommendations and 4 `metadata-only` recommendations;
- malformed bytes are rejected;
- risk families are `preservedActiveX`, `preservedMacro`, `preservedOther`, and `preservedSignature`;
- compact JSON omits `inputSha256`, workbook `bytes`, and timing sample fields.

## Confidence

High for local proof packaging. Medium for release wording because public/release owners still need to accept generated signed and unknown-part cases or replace them with public binaries, and performance owners still need release-environment latency evidence.

## Fold-in decision

Folded into the benchmark proof harness only. This is not a new open surface.

## Next question

Should safe-open compact reports also stay generated on demand until the same artifact storage/privacy/canonicalization policy exists for package-action compact reports?
