# Package Action Owner Acceptance Checklist

## Question

Can package-action owner acceptance be reduced to exact checkboxes covering edge fixtures, unsupported features, streaming wording, provenance, and compact report publication?

## Hypothesis

Yes. The current package-action proof and release index already carry the needed evidence and blockers. The next useful artifact is an owner checklist, not another package mutation surface.

## External sources checked

- Microsoft Excel calculation chain metadata: https://support.microsoft.com/en-us/office/excel-calculation-chain-metadata-6e1b5819-6abd-4e94-bff5-838d4c576e01
- Microsoft Open Packaging Conventions fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl tutorial preservation boundary: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

Auditable package-part mutation is the second release handoff. Without exact owner checkboxes, local per-part evidence can be overread as semantic support for unsupported Excel features, full streaming parity, or provenance.

## Probe/implementation

- Ran a local Bun probe over `runPackageActionProof`, `packageActionCompactReleaseReport`, and `runReleaseProofIndex`.
- Confirmed package-action proof metrics and missing `readyWhen` gates.
- Updated `research/experiments/syntheses/2026-05-owner-handoff.md` with package-action acceptance checkboxes.
- Updated `research/experiments/syntheses/2026-05-release-claim-board.md` with the same owner-ready boundary.

## Results

Current proof input:

| Field | Value |
| --- | --- |
| Proof cases | 8 |
| Public fixture cases | 2 |
| Generated workbook cases | 2 |
| Generated edge-package cases | 4 |
| Generated/disclosed cases | `docprops-passthrough`, `regenerate-existing-sheet`, `add-sheet-part`, `calc-chain-drop`, `signature-invalidation-drop`, `unknown-part-error` |
| Action classes | `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1` |
| Source graph evidence everywhere | `true` |
| Journal package issues everywhere | `true` |
| Post-write audit failures | `unknown-part-error` |
| Representative streaming proof cases | 1 |
| Headline claim allowed | `false` |
| Release gate | `blocked-by-publication-policy` |

Acceptance checklist:

- Product accepts disclosed generated edge packages for docProps passthrough, calc-chain drop, signature invalidation drop, and unknown-part error, or replaces them with public binary fixtures.
- Correctness approves unsupported-feature boundaries for signatures, calc chain, chart/drawing sidecars, macros/ActiveX, and unknown parts.
- Correctness keeps journal/package issue compatibility as part of the claim.
- Performance accepts one representative streaming dirty-sheet proof as sufficient for narrow wording, or expands streaming variants before broader streaming claims.
- Release approves local-proof wording that excludes SLSA, in-toto, signed provenance, third-party attestation, and tamper-evident storage.
- Release keeps compact report digests unpublished until storage, privacy filtering, canonicalization, and verification expectations exist.

## Confidence

High that the checklist reflects current proof. Medium that the claim can be released, because owner approval remains outside research.

## Fold-in decision

Promote to topic synthesis only. Keep `package-action-proof` blocked until product, correctness, performance, and release owners complete the checklist. Do not add package mutation surfaces.

## Next question

Can the owner handoff be machine-checked by making a small Markdown-only acceptance table in the release claim board the single source of truth, or would that duplicate `release-proof-index`?
