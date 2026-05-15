# Package Action Claim Proof Refresh

## Question

Can the second-ranked "auditable package-part mutation" claim be refreshed from existing package-action proof evidence without adding a new mutation surface?

## Hypothesis

Yes. The tracked package-action proof harness already covers the release vocabulary: `passthrough`, `regenerate`, `add`, `drop`, and `error`. The useful work is to rerun it, verify SDK/CLI/API/MCP evidence paths, and keep the provenance boundary honest.

## External sources checked

- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft `System.IO.Packaging`: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging
- openpyxl tutorial and preservation warnings: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/
- in-toto attestations: https://github.com/in-toto/attestation

## Why this matters to Ascend

Ascend's strongest correctness/product differentiation is not only that it writes XLSX, but that it can explain what happened to each relevant package part. This claim needs evidence that is strong enough for release copy while refusing stronger provenance or Excel-certification language.

## Probe/implementation

- Inspected `research/experiments/syntheses/2026-05-package-action-proof-report.md`.
- Inspected the tracked harness and tests:
  - `fixtures/benchmarks/package-action-proof.ts`
  - `fixtures/benchmarks/package-action-proof.test.ts`
- Reran:
  - `bun run fixtures/benchmarks/package-action-proof.ts`
  - `bun run fixtures/benchmarks/package-action-proof.ts --json`
- Reran validation:
  - `bun test fixtures/benchmarks/package-action-proof.test.ts`
  - `bun test packages/sdk/src/agent-workflow.test.ts -t "package action|package graph|journalSummary|compact commit"`
  - `bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow"`
  - `bun test apps/api/api.test.ts -t "plan and commit endpoints provide the safe write workflow"`
  - `bun test apps/mcp/src/index.test.ts -t "package action proof evidence"`
- Updated the package-action proof report with fresh timings and validation status.

## Results

The harness still covers all five action kinds across eight cases:

| Case | Commit actions | Journal package issues | Proof issues | Post-write audit |
| --- | --- | ---: | ---: | --- |
| docprops-passthrough | passthrough=4, regenerate=4, add=0, drop=0, error=0 | 1 | 0 | passed |
| regenerate-existing-sheet | passthrough=3, regenerate=5, add=0, drop=0, error=0 | 1 | 0 | passed |
| add-sheet-part | passthrough=3, regenerate=5, add=1, drop=0, error=0 | 1 | 0 | passed |
| calc-chain-drop | passthrough=0, regenerate=5, add=0, drop=1, error=0 | 1 | 0 | passed |
| signature-invalidation-drop | passthrough=1, regenerate=4, add=0, drop=2, error=0 | 1 | 0 | passed |
| macro-passthrough | passthrough=6, regenerate=5, add=1, drop=0, error=0 | 1 | 0 | passed |
| chart-sidecar-accounting | passthrough=8, regenerate=6, add=1, drop=0, error=0 | 1 | 0 | passed |
| unknown-part-error | passthrough=2, regenerate=4, add=0, drop=0, error=1 | 1 | 1 | needs review |

Combined commit actions: `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1`.

SDK, CLI, API, and MCP validations passed. No new mutation surface was added.

## Confidence

High for guarded release wording around local package-action evidence. Medium for publication packaging because synthetic edge cases still need an explicit publication policy, and the proof must not be mistaken for signed provenance or Excel semantic certification.

## Fold-in decision

Promote to product/correctness proof packaging only. Do not add a new mutation surface. Keep the proof beside safe-open as release evidence, with boundaries around signatures, chart XML regeneration, and provenance.

## Next question

Can the release proof index point to the refreshed safe-open and package-action reports by digest without embedding bulky generated JSON or implying tamper-evident attestation?
