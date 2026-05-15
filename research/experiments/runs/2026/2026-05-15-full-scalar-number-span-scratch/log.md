# Full scalar number span scratch

## Question

Can the XLSX full-scalar byte parser avoid repeated short numeric-span array allocation without changing read semantics?

## Hypothesis

Yes. The parser only needs a transient array when adjacent numeric cells become a dense span. Reusing a per-sheet scratch array and resetting `length` should avoid allocating a new two-value array for each detected span while preserving output cells.

## External sources checked

- MDN `Array` documentation describes the `length` property and `push()` updating array contents: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array
- V8 elements-kind documentation explains why stable array element kinds and array operations matter in hot JavaScript paths: https://v8.dev/blog/elements-kinds
- Microsoft Open XML formulas/read overview frames worksheet XML parsing as a reader responsibility: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas

## Why this matters to Ascend

Real-world XLSX performance is a North Star claim. The full-scalar byte parser is a hot read path for simple dense sheets; reducing allocation pressure there supports performance without changing workbook semantics.

## Probe/implementation

Updated the reader parse context:

- `packages/io-xlsx/src/reader/index.ts` passes a per-sheet `fullScalarNumberSpanScratch: []`.
- `packages/io-xlsx/src/reader/sheet.ts` reuses that scratch array when starting adjacent numeric spans, resets `length`, and pushes the pending/current values.

Committed production change:

```text
0f16dfde perf(io-xlsx): reuse numeric span scratch during sheet read
```

Validation:

```bash
bun test packages/io-xlsx/src/reader/reader.test.ts -t "full scalar byte parser|full mode preserves simple numeric"
bunx biome check packages/io-xlsx/src/reader/index.ts packages/io-xlsx/src/reader/sheet.ts
bunx tsc --build
bun run test:changed
```

## Results

- Reader targeted tests passed: 3 tests, 20 assertions.
- Biome passed.
- TypeScript build passed.
- `bun run test:changed` passed: 4378 pass, 1 skip, 0 fail across 168 files.
- This is a tiny allocation hygiene change, not a new performance claim. No benchmark threshold was changed.

## Confidence

Medium. The semantic tests cover the full-scalar parser path, but this log does not include heap profiling or a statistically meaningful benchmark. Treat as a safe local optimization, not a release performance proof.

## Fold-in decision

Promote to performance loop as scoped allocation hygiene. Do not promote as a columnar/reader performance claim until a benchmark or heap profile proves impact.

## Next question

Should the performance loop add a heap-allocation probe for full-scalar dense numeric XLSX reads before making further parser allocation changes?
