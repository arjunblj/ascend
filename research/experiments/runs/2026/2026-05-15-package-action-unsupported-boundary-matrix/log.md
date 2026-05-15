# Package Action Unsupported Boundary Matrix

## Question

Can the package-action `unsupported-feature-boundary` gate be turned into an approval-ready correctness matrix without adding new mutation surfaces?

## Hypothesis

Yes. The package-action proof already covers the hard unsupported-feature classes. The missing work is claim wording: distinguish package accounting from semantic support for signatures, calc chains, chart XML, macros/ActiveX, unknown parts, and streaming parity.

## External sources checked

- Microsoft OPC fundamentals and digital signatures: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft OPC article on signed parts and relationships: https://learn.microsoft.com/en-us/archive/msdn-magazine/2007/august/opc-a-new-standard-for-packaging-your-data
- Microsoft Excel calculation chain metadata: https://support.microsoft.com/en-us/office/excel-calculation-chain-metadata-6e1b5819-6abd-4e94-bff5-838d4c576e01
- OOXML unknown relationships reference: https://ooxml.info/docs/9/9.1/9.1.7/
- SheetJS local file/write docs: https://docs.sheetjs.com/docs/demos/local/file
- SheetJS write options reference mirror: https://github.com/observablehq/xlsx
- openpyxl tutorial and workbook object docs: https://openpyxl.readthedocs.io/en/3.1/tutorial.html

## Why this matters to Ascend

Auditable package-part mutation is the second release-claim handoff. Without a correctness-approved boundary matrix, per-part evidence can be misread as semantic support for unsupported Excel features or as signed provenance.

## Probe/implementation

- Inspected `fixtures/benchmarks/package-action-proof.ts`, `fixtures/benchmarks/release-proof-index.ts`, and the release claim board.
- Ran `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json`.
- Ran `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json`.
- Ran `bun test fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/release-proof-index.test.ts`.
- Updated `research/experiments/syntheses/2026-05-release-claim-board.md` with an unsupported-feature boundary matrix.

## Results

Package-action proof rerun:

| Metric | Value |
| --- | ---: |
| Cases | 8 |
| Public fixture cases | 2 |
| Generated workbook cases | 2 |
| Generated edge-package cases | 4 |
| Passthrough actions | 27 |
| Regenerate actions | 38 |
| Add actions | 3 |
| Drop actions | 3 |
| Error actions | 1 |
| Source graph evidence in every case | yes |
| Package journal issue in every case | yes |
| Post-write audit failures | `unknown-part-error` only |
| Streaming proof cases | 1 |

Boundary matrix:

| Feature | Allowed claim | Forbidden claim |
| --- | --- | --- |
| Digital signatures | detect and report invalidation/drop evidence | preserve, verify, re-sign, attest |
| Calc chain | report drop/regeneration decisions for unsafe cached calculation order | prove Excel recalculation equivalence |
| Chart/drawing sidecars | account for sidecars separately from regenerated parts | claim chart XML byte passthrough or full semantic chart support |
| Macros/ActiveX | record package preservation and safe-open review routing | claim scanning, sandboxing, execution safety |
| Unknown parts | fail closed with explicit unknown-part error | preserve or understand arbitrary unknown parts |
| Streaming writer | one representative streaming proof exists | claim full streaming parity |

## Confidence

High that the boundary matrix matches current proof. Medium that it is sufficient for release copy; a correctness owner still needs to approve the allowed/forbidden wording.

## Fold-in decision

Promote to topic synthesis and owner-loop handoff only. Do not change package-action surfaces. Keep `unsupported-feature-boundary` missing in the release proof index until correctness explicitly approves the boundary matrix.

## Next question

Can the package-action `provenance-boundary` gate be made approval-ready with a release-owner matrix that separates local proof digests from signed provenance, SLSA, in-toto, and third-party attestation?
