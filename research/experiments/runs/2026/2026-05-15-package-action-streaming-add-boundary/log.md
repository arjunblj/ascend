# Package Action Streaming Add Boundary

## Question

Can the package-action streaming proof cover an `add` action without weakening the release boundary that still blocks full streaming-parity wording?

## Hypothesis

`add-sheet-part` should be safe to include as a second representative streaming probe because it exercises a generated workbook topology mutation and should produce a clean streaming proof with passthrough byte equality. `calc-chain-drop` should not be promoted unless the streaming proof actually emits the expected `drop` action for `xl/calcChain.xml`.

## External sources checked

- Node.js stream documentation: https://nodejs.org/api/stream.html
- Bun incremental file sink documentation: https://bun.com/docs/guides/write-file/filesink
- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Excelize streaming write documentation: https://xuri.me/excelize/en/stream.html

## Why this matters to Ascend

Auditable package-part mutation is one of Ascend's top release-claim candidates. Streaming writer evidence can support performance-oriented wording only if it stays tied to concrete package actions and explicit missing classes; otherwise a narrow probe can accidentally become a broad "streaming parity" claim.

## Probe/implementation

- Inspected `fixtures/benchmarks/package-action-proof.ts` and the release proof index owner-handoff expectations.
- Added `streamingProbe: true` to `add-sheet-part`.
- Tried the same for `calc-chain-drop`, but the streaming proof did not emit the expected `drop` action for `xl/calcChain.xml`, so that attempted promotion was rejected.
- Updated package-action proof tests and release-proof-index tests so the streaming matrix now reports covered action kinds `passthrough`, `regenerate`, and `add`, with `drop` and `error` still missing.
- Tightened release wording from "one representative proof" to "representative streaming proofs" while still forbidding full parity, `drop`/`error` streaming behavior, and public macro/chart streaming coverage claims.

## Results

Validation command:

```bash
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
```

Key result:

- `streamingProofCases=2`
- `streamingRegenerateParts=1`
- `docprops-passthrough` streaming proof: `regeneratePartCount=1`, `passthroughBytesEqualCount=3`, `issueCount=0`
- `add-sheet-part` streaming proof: `regeneratePartCount=0`, `passthroughBytesEqualCount=3`, `issueCount=0`
- Combined package-action proof remains `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1`
- Release gate remains `blocked-by-publication-policy`; `streaming-matrix-boundary` remains missing.

Targeted validation:

```bash
bun test fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/release-proof-index.test.ts
bunx biome check apps/api/src/server.ts apps/api/src/server.test.ts fixtures/benchmarks/package-action-proof.ts fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bunx tsc --build
```

All passed.

## Confidence

Medium-high for the narrow proof expansion: the harness now has two passing streaming cases and machine-visible missing action classes. Medium-low for broader streaming claims: `drop`, `error`, macro, and chart streaming behavior remain unproved.

## Fold-in decision

Fold into the performance/release proof harness only. Do not promote a new SDK, CLI, API, or MCP surface. Keep `streaming-matrix-boundary` owner-owned until performance either accepts this narrow wording or expands the matrix to `drop`, `error`, and public macro/chart cases.

## Next question

Can release owner handoffs make the remaining `drop`/`error` streaming gap more actionable without adding more production surface area?
