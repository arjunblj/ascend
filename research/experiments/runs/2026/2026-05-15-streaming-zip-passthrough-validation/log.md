# Streaming ZIP Passthrough Validation

## Question

Can streaming XLSX writes preserve unchanged source ZIP members as compressed passthrough entries during dirty-sheet edits, instead of inflating and recompressing preserved parts?

## Hypothesis

Yes. The non-streaming writer already uses source-archive passthrough for preserved parts. The streaming writer can carry the same source archive and passthrough path set into `StreamingZipBuilder`, then reuse matching compressed bytes for non-streamed preserved entries while still streaming regenerated worksheet XML.

## External sources checked

- TypeScript `exactOptionalPropertyTypes` docs: https://www.typescriptlang.org/tsconfig/exactOptionalPropertyTypes.html
- TypeScript 4.4 release notes for exact optional property semantics: https://www.typescriptlang.org/docs/handbook/release-notes/typescript-4-4.html
- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

The North Star includes preservation-first XLSX and real-world performance. Dirty-sheet edits should avoid touching unchanged package parts, and the streaming writer should not regress from the normal writer's compressed-passthrough behavior. This is directly relevant to auditable package-part mutation and practical writer latency because unchanged shared strings, metadata, and sidecars can remain byte-stable without decompression/recompression work.

## Probe/implementation

- Inspected the dirty writer changes in `packages/io-xlsx/src/writer/index.ts`, `packages/io-xlsx/src/writer/zip.ts`, and `packages/io-xlsx/src/writer/writer.test.ts`.
- Finished the TypeScript shape so `createZipStreaming` passes a proper `StreamingZipBuilder` options object under `exactOptionalPropertyTypes`.
- Extended `StreamingZipBuilder.addEntry` to call the existing compressed-passthrough helper before recompressing normal entries.
- Added a streaming dirty-sheet test that:
  - mutates only `sheet1.xml`;
  - asserts `xl/sharedStrings.xml` is not read as text;
  - asserts the written compressed shared-strings bytes equal the source compressed bytes;
  - asserts the changed worksheet CRC differs from the source.

## Results

Validation:

```bash
bun test packages/io-xlsx/src/writer/writer.test.ts -t "passes through preserved source ZIP parts"
bunx biome check packages/io-xlsx/src/writer/index.ts packages/io-xlsx/src/writer/zip.ts packages/io-xlsx/src/writer/writer.test.ts
bunx tsc --build
bun run test:changed
```

Results:

- Focused writer passthrough tests: 2 pass, 0 fail.
- TypeScript build: pass.
- Biome check: pass.
- Changed-test/full suite rerun: 5065 pass, 1 skip, 0 fail.

## Confidence

High for the scoped fold-in. The implementation reuses the existing passthrough eligibility helper and has parity tests for non-streaming and streaming dirty-sheet edits. Medium for broad performance impact because this validates byte preservation and avoided text reads, not a latency benchmark over large real workbooks.

## Fold-in decision

Promote to the performance/correctness loop as a scoped writer improvement. This is not a new product surface and does not change claim wording by itself. It strengthens the evidence behind preservation-first dirty writes and package-part accounting.

## Next question

Should the package-action proof harness add a streaming-writer variant so release proof can show passthrough/regenerate/add/drop/error behavior for both normal and streaming writer paths?
