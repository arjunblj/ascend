# Package Action Release Index Blockers

## Question

Can the auditable package-part mutation artifact receive the same release-readiness treatment as safe-open while preserving no-attestation and chart-byte-boundary language?

## Hypothesis

Yes. The package-action proof is already one of the two release-index artifacts. The useful work is to rerun the proof, align the synthesis with the new release-index publication blockers, and avoid adding any mutation surface.

## External sources checked

- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- `System.IO.Packaging` package concepts and signatures: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging
- openpyxl tutorial preservation warning: https://openpyxl.pages.heptapod.net/openpyxl/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/
- in-toto Attestation Framework: https://github.com/in-toto/attestation

## Why this matters to Ascend

The claim is product-shaped: "auditable package-part mutation," not merely "writes XLSX." It helps Ascend be trustworthy when editing real workbooks because package parts can be accounted for as `passthrough`, `regenerate`, `add`, `drop`, or `error` while preserving boundaries around signatures, unsupported features, and provenance.

## Probe/implementation

No production mutation code changed. Reran existing proof artifacts:

```bash
bun run fixtures/benchmarks/package-action-proof.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
bun test fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/release-proof-index.test.ts
```

Updated `research/experiments/syntheses/2026-05-package-action-proof-report.md` with the latest package-action proof table and release-index blocker status.

## Results

Package-action proof rerun:

- 8 proof cases.
- Combined commit actions: passthrough=27, regenerate=38, add=3, drop=3, error=1.
- Source graph evidence in every case.
- Package-preservation journal issue in every case.
- Unknown-part case reports one proof issue and needs review.

Release proof index result:

- Artifact: `package-action-proof`.
- Command: `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json`.
- Publication status: `needs-release-packaging`.
- Publication blockers:
  - synthetic edge packages must stay disclosed unless replaced by public binary fixtures;
  - proof is local evidence, not signed provenance or third-party attestation.
- Stable shape digest: `b9758496346c97920c80ba08b6632315708a6d6cc770927695337e729554dbb0`.

The chart boundary remains: do not claim chart byte passthrough. Claim only per-part accounting; chart XML can regenerate while drawing sidecars pass through.

Validation passed: 6 tests, 0 failures across the package-action proof and release-proof-index harnesses.

## Confidence

High for the local proof and release-index alignment. Medium for public release wording until product decides whether code-generated edge packages are acceptable as disclosed proof fixtures.

## Fold-in decision

Promote to correctness/product proof packaging. Do not add mutation surfaces. Keep the artifact as local evidence and do not imply signed provenance, SLSA, in-toto, or Excel semantic certification.

## Next question

Can the top-two release proof artifacts be summarized as next-loop owner prompts with no additional implementation in this block?
