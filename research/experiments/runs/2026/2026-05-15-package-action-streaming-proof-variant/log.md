# Package Action Streaming Proof Variant

## Question

Should the package-action proof harness include a streaming writer variant so release evidence can show that package action accounting still works when worksheet XML is generated through the streaming writer path?

## Hypothesis

Yes, narrowly. The proof should not add a new SDK/CLI/API/MCP surface, but the benchmark harness can run one representative streaming writer probe and record whether regenerated worksheet parts are marked as streaming while preserved source parts still retain digest evidence.

## External sources checked

- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- SheetJS write options and data-preservation framing: https://docs.sheetjs.com/docs/api/write-options/
- SheetJS streaming export API: https://docs.sheetjs.com/docs/api/stream/

## Why this matters to Ascend

Auditable package-part mutation is one of the top release claims. Ascend recently gained streaming ZIP passthrough parity for dirty-sheet edits, but the release proof harness still only exercised the buffered SDK commit path. If streaming write summaries lose the streaming marker or silently reclassify package parts, owner loops cannot tell whether the performance path preserves the same part-action semantics.

## Probe/implementation

- Inspected `fixtures/benchmarks/package-action-proof.ts`, `packages/io-xlsx/src/writer/index.ts`, and `packages/io-xlsx/src/writer/plan.ts`.
- Added a benchmark-only `streamingProof` probe to the `docprops-passthrough` package-action case.
- The probe:
  - reads the same generated docProps package;
  - applies the same operation locally;
  - runs `summarizePlannedWrite(..., { streaming: true })`;
  - writes with `writeXlsxStreaming`;
  - builds a digest-backed `ascend-package-action-proof` from the streaming summary and output bytes.
- Fixed summary-only streaming plan recording so `summarizePlannedWrite` preserves `streaming: true` for streaming-generated worksheet parts.
- Surfaced `streamingProofCases` and `streamingRegenerateParts` in the release proof index summary for the package-action artifact.

## Results

Current streaming proof for `docprops-passthrough`:

| Metric | Value |
| --- | --- |
| streaming expected actions present | true |
| streaming regenerated parts | `xl/worksheets/sheet1.xml` |
| streaming passthrough byte-equal parts | 3 |
| streaming output digest count | 8 |
| streaming proof issues | 0 |

Current release proof index package-action stable shape digest: `9abebf576651551f58e00ccf8469d099b2c06dacd48391fe581a24e51a1e0afd`.

Validation:

```bash
bun test fixtures/benchmarks/package-action-proof.test.ts packages/io-xlsx/src/writer/writer.test.ts -t "streaming write plan emits worksheet XML through streamingBuild|package action proof"
bunx biome check packages/io-xlsx/src/writer/index.ts packages/io-xlsx/src/writer/plan.ts fixtures/benchmarks/package-action-proof.ts fixtures/benchmarks/package-action-proof.test.ts
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

## Confidence

High for the narrow proof: the harness now catches summary-only loss of streaming metadata and validates digest-backed package action evidence for one streaming dirty-sheet case. Medium for broad release claims because only one representative streaming case is covered; the owner loop should not claim every package-action scenario has streaming parity yet.

## Fold-in decision

Fold into the performance/correctness proof harness. Do not promote a new product surface. Use this as supporting evidence for auditable package-part mutation and streaming writer preservation, with boundary language that streaming proof currently covers one representative docProps dirty-sheet case.

## Next question

Should the release proof index readiness gates require streaming parity for all package-action scenarios, or is one representative streaming proof enough until a performance owner requests full matrix coverage?
