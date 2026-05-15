# Package Action Release Proof Rerun

## Question

Can auditable package-part mutation receive a current release-proof rerun and compact report now that the tracked worktree is clean?

## Hypothesis

Yes. The tracked `package-action-proof` harness should produce current evidence for all five action kinds without changing writer behavior or adding a new product surface.

## External sources checked

- Microsoft Open Packaging Conventions fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft `System.IO.Packaging`: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging
- openpyxl tutorial and preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options and unsupported-feature boundary: https://docs.sheetjs.com/docs/api/write-options/
- in-toto attestation framework: https://github.com/in-toto/attestation

## Why this matters to Ascend

This is the second-ranked product-shaped release claim after safe unknown workbook opening. It is valuable only if it stays precise: local package-part accounting, not signed provenance or a blanket Excel-compatible writer claim.

## Probe/implementation

- Confirmed the tracked worktree no longer had dirty writer files.
- Reran:

```bash
bun run fixtures/benchmarks/package-action-proof.ts
bun test fixtures/benchmarks/package-action-proof.test.ts
```

- Created `research/experiments/syntheses/2026-05-package-action-proof-report.md`.
- Updated the experiment index and release claim board.
- Did not change writer behavior or product surfaces.

## Results

Combined commit actions:

- `passthrough=27`
- `regenerate=38`
- `add=3`
- `drop=3`
- `error=1`

Current local boundaries:

- Chart/drawing content is accounted for per part; chart XML is regenerated while drawing sidecars pass through.
- Unknown package parts produce explicit review-required error evidence.
- The report is local evidence and does not claim SLSA, in-toto, signed provenance, or Excel recalculation equivalence.

Validation already green in this loop:

```bash
bun test fixtures/benchmarks/package-action-proof.test.ts
bunx tsc --build
bun run test:changed
```

## Confidence

High for guarded local proof language. Medium for public publication readiness until product decides whether synthetic signed/calc-chain/unknown fixtures should become durable public binary fixtures.

## Fold-in decision

Promote to product/correctness proof packaging as the second release-proof artifact. Do not add another SDK/CLI/API/MCP surface.

## Next question

Can retained viewport patch history be rerun and ranked against token-bounded agent view without promoting another release headline?
