# Fast Zip Filtered Worksheet Compression

## Question

Should fast XLSX ZIP compression use `Z_FILTERED` for worksheet XML below the huge-sheet threshold?

## Hypothesis

Yes. Worksheet XML is repetitive tag-heavy text, so zlib's filtered strategy can reduce fast-profile deflate time without increasing output size.

## External sources checked

- zlib manual: https://zlib.net/manual.html
- Bun benchmarking and profiling documentation: https://bun.sh/docs/project/benchmarking
- hyperfine benchmark tool: https://github.com/sharkdp/hyperfine

## Why this matters to Ascend

Write performance is part of the real-world spreadsheet platform claim, but compression changes need evidence because file size regressions can erase any wall-time gain.

## Probe/implementation

Implementation changes fast ZIP compression so regular worksheet XML uses `{ level: 1, strategy: Z_FILTERED }`; huge worksheet XML already used the filtered strategy.

Local zlib probe:

- generated worksheet-shaped XML with 10,000 rows x 20 numeric cells.
- compressed it 8 times with `level: 1`.
- compressed it 8 times with `level: 1, strategy: Z_FILTERED`.

## Results

Input XML bytes: 3,992,151.

| Strategy | Median | Min | Max | Output bytes |
| --- | ---: | ---: | ---: | ---: |
| level 1 default | 7.656 ms | 7.206 ms | 8.830 ms | 1,032,753 |
| level 1 filtered | 7.059 ms | 6.907 ms | 7.713 ms | 1,032,753 |

The local probe showed an 8.5% median deflate improvement with equal compressed size.

Validation:

- `bun test packages/io-xlsx/src/writer/writer.test.ts`
- `bunx biome check packages/io-xlsx/src/writer/zip.ts`
- `bunx tsc --build`
- `bun run test:changed`

## Confidence

Medium-low. The probe supports the direction, and writer tests protect validity, but this is a synthetic zlib-only benchmark. A full XLSX write-phase benchmark should own any broader performance claim.

## Fold-in decision

Promote as a tiny performance-loop fold-in. Do not promote a release performance threshold.

## Next question

Can the XLSX write-phase harness expose compression profile and worksheet strategy attribution without expanding release benchmarks?
