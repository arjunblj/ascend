# 2026-05 Package Action Proof Report

Date: 2026-05-15

## Claim

Ascend can explain XLSX writes with local package-part action evidence: `passthrough`, `regenerate`, `add`, `drop`, and `error`.

## Claim Wording That Is Safe Today

Ascend can produce a local package-action proof for representative workbook mutations, showing which package parts were passed through, regenerated, added, dropped, or rejected with review-required errors. This is local evidence, not signed provenance, SLSA, in-toto, or Excel semantic certification.

## External Contrast

- [Microsoft OPC fundamentals](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview) define packages as parts plus relationships and describe package signatures as validation evidence over signed content.
- [Microsoft `System.IO.Packaging`](https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging) exposes package parts, relationships, and signatures as first-class concepts, matching Ascend's per-part proof boundary.
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/tutorial.html) documents `keep_vba` and warns that unsupported workbook objects may be lost on save.
- [SheetJS write options](https://docs.sheetjs.com/docs/api/write-options/) state that undocumented features may not serialize, which is a useful contrast for Ascend's explicit package-action accounting.
- [in-toto attestations](https://github.com/in-toto/attestation) are real supply-chain provenance artifacts; Ascend's package-action proof must not imply that level of signed attestation.

## Proof Bundle Status

| Required proof | Current evidence | Status |
| --- | --- | --- |
| Fixture mix | Synthetic docProps, new workbook edits, calc-chain, signature, unknown package part; public macro and chart workbooks | Covered for local proof; synthetic edge cases may still need durable public binaries before external publication |
| Action vocabulary | Combined commit evidence covers `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1` | Covered |
| SDK evidence shape | Tracked `fixtures/benchmarks/package-action-proof.ts` harness uses existing SDK plan/commit and package-action proof helpers | Covered |
| Validation gate | Harness test, prior full `test:changed`, typecheck, and Biome on changed TypeScript files | Covered in current loop |
| Competitor contrast | OPC, openpyxl, SheetJS, in-toto boundary | Covered |
| Honest boundary | Chart XML regenerates while drawing sidecars pass through; proof is local package evidence, not signed provenance or Excel recalc equivalence | Covered |

## Fresh Local Probe

Probe command:

```bash
bun run fixtures/benchmarks/package-action-proof.ts
```

| Case | Fixture | Input bytes | Output bytes | Commit actions | Digest pairs | Issues | Proof JSON bytes | Proof ms | Post-write audits | Example actions |
| --- | --- | ---: | ---: | --- | ---: | ---: | ---: | ---: | --- | --- |
| docprops-passthrough | synthetic docProps package | 2286 | 3485 | passthrough=4, regenerate=4, add=0, drop=0, error=0 | 8 | 0 | 5155 | 0.224 | passed | passthrough workbook; regenerate sheet |
| regenerate-existing-sheet | new Ascend workbook | 4624 | 4707 | passthrough=3, regenerate=5, add=0, drop=0, error=0 | 8 | 0 | 5132 | 0.074 | passed | passthrough workbook; regenerate styles |
| add-sheet-part | new Ascend workbook | 4624 | 4512 | passthrough=3, regenerate=5, add=1, drop=0, error=0 | 8 | 0 | 5610 | 0.071 | passed | add worksheet part |
| calc-chain-drop | synthetic calcChain package | 1776 | 2365 | passthrough=0, regenerate=5, add=0, drop=1, error=0 | 5 | 0 | 3760 | 0.078 | passed | drop calcChain |
| signature-invalidation-drop | synthetic digital-signature package | 2253 | 2058 | passthrough=1, regenerate=4, add=0, drop=2, error=0 | 5 | 0 | 4165 | 0.087 | passed | drop signature parts |
| macro-passthrough | `fixtures/xlsx/calamine/vba.xlsm` | 12752 | 12175 | passthrough=6, regenerate=5, add=1, drop=0, error=0 | 11 | 0 | 7359 | 0.199 | passed | passthrough VBA project |
| chart-sidecar-accounting | `fixtures/xlsx/poi/WithChart.xlsx` | 10138 | 10899 | passthrough=8, regenerate=6, add=1, drop=0, error=0 | 14 | 0 | 9067 | 0.201 | passed | passthrough drawing; regenerate chart/styles |
| unknown-part-error | synthetic unknown package part | 1692 | 2315 | passthrough=2, regenerate=4, add=0, drop=0, error=1 | 7 | 1 | 4629 | 0.081 | needs review | error unknown custom part |

## Interpretation

- The proof is strong enough for guarded release language around local package-part accounting.
- The proof should not say "chart byte passthrough." It should say chart/drawing package content is accounted for with per-part actions.
- The unknown-part case is intentionally not clean: it demonstrates explicit review-required evidence rather than silent preservation claims.
- Timing and JSON-size values are evidence for report shape, not performance thresholds.

## Fold-In Recommendation

Promote as the second release-proof artifact beside safe-open. Do not add a new mutation surface. The next product step is publication packaging: decide where the Markdown/JSON report lives and whether synthetic edge cases need public binary fixtures.
