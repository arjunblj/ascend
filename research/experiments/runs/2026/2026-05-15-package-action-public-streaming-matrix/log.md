# Package Action Public Streaming Matrix

## Question

Can the package-action proof harness cover public macro and chart fixtures through the streaming writer path without promoting a full streaming parity claim?

## Hypothesis

If the existing streaming writer can produce package-action proof for the public macro and chart cases with zero proof issues, then the `streaming-matrix-boundary` blocker can be narrowed from "macro/chart unproven" to "generated edge/error cases remain outside parity." The release claim should still stay below full streaming parity.

## External sources checked

- Quadratic product/docs for QSS contrast around AI/code spreadsheet UX: https://docs.quadratichq.com/
- Quadratic import note that Excel files must be converted to CSV for drag-and-drop flows: https://docs.quadratichq.com/import-data/drag-and-drop-.parquet
- OOXML package parts and relationships overview: https://ooxml.info/docs/8/8.2/
- Microsoft Open XML SDK spreadsheet package concepts: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document

## Why this matters to Ascend

The QSS-leapfrog package-action claim is strongest when Ascend can prove what happened to workbook package parts after mutation. QSS appears stronger as an AI/code spreadsheet product surface, but Ascend can be better as a preservation/proof runtime if package-action evidence covers real public workbook feature families without overclaiming.

## Probe/implementation

Inspected `fixtures/benchmarks/package-action-proof.ts` and `fixtures/benchmarks/release-proof-index.ts`. The existing proof harness already covered public macro and chart fixtures in the standard writer path, but only `docprops-passthrough`, `add-sheet-part`, and `calc-chain-drop` had streaming proof probes.

Ran a local throwaway streaming probe for:

- `macro-passthrough` using `fixtures/xlsx/calamine/vba.xlsm`
- `chart-sidecar-accounting` using `fixtures/xlsx/poi/WithChart.xlsx`

Both probes used `readXlsx`, `applyOperations`, `summarizePlannedWrite`, `writeXlsxStreaming`, and `createPackageActionProof` over the same operations as the committed package-action proof harness.

Folded the successful result into the proof harness:

- enabled `streamingProbe: true` for `macro-passthrough`
- enabled `streamingProbe: true` for `chart-sidecar-accounting`
- updated compact proof and release-index expectations from 3 to 5 streaming proof cases
- updated release wording so generated edge/error cases remain the streaming boundary

## Results

Local probe output:

| Case | Result |
| --- | --- |
| `macro-passthrough` | `issues=0`, `passthrough=5`, `regenerate=6`, expected macro passthrough present |
| `chart-sidecar-accounting` | `issues=0`, `passthrough=7`, `regenerate=7`, expected chart regenerate and drawing passthrough present |

Committed proof output now reports:

- `streamingProofCases=5`
- `streamingRegenerateParts=4`
- covered action kinds: `passthrough`, `regenerate`, `add`, `drop`
- missing action kind: `error`
- public non-streaming cases: none
- generated non-streaming cases: `regenerate-existing-sheet`, `signature-invalidation-drop`, `unknown-part-error`

Validation:

- `bun test fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/release-proof-index.test.ts --timeout 30000`
- `bunx biome check fixtures/benchmarks/package-action-proof.ts fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts`
- `bunx tsc --build`
- `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`
- `bun run test:changed`

## Confidence

High that public macro/chart package-action streaming evidence is now covered by the proof harness. Medium for release wording because generated edge/error streaming cases remain outside the proof boundary and owner approval is still required.

## Fold-in decision

Promote as a proof-harness fold-in only. Do not add SDK, CLI, API, or MCP surfaces. Keep `streaming-matrix-boundary` owner-owned and forbid full streaming parity, generated edge/error streaming, signed provenance, and semantic support claims for unsupported workbook features.

## Next question

Should the performance loop attempt a generated `error` streaming proof, or should release wording permanently accept the current five-case public streaming matrix as representative evidence only?
