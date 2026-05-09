# Benchmarks

## Synthetic suite (CI)

- **Command:** `bun bench` (runs `fixtures/benchmarks/run.ts`).
- **Smoke set in CI:** `--set smoke --repeat 3 --ci` then `fixtures/benchmarks/check-targets.ts`.
- **Throughput gates** (best scenario per category): see [`fixtures/benchmarks/targets.ts`](fixtures/benchmarks/targets.ts) — read ≥ 3M cells/s, write ≥ 1.5M/s, calc ≥ 500K/s.

## Microbenchmarks

```bash
bun run fixtures/benchmarks/micro.ts
```

## Competitive (local)

| Script | Purpose |
|--------|---------|
| `bun run bench:competitive` | Ascend vs HyperFormula (calc scenarios) |
| `bun run bench:competitive-io` | Ascend vs SheetJS / ExcelJS read+write, plus installed cross-language write/read runners such as XlsxWriter, OpenPyXL, PyExcelerate, fastexcel, Calamine, Excelize, Apache POI, and ClosedXML |
| `bun run bench:upstream-profiles` | Ascend vs OSS libraries on benchmark shapes published by those projects |
| `bun run bench:upstream-real` | Ascend vs OSS libraries on published real-workbook benchmark files |
| `bun run bench:competitive-memory` | Rough RSS / heap snapshots after workloads |

`bench:competitive-io` accepts `--category read|write|all`, `--competitor js|external|python|all`, `--runner-manifest` for generated read runners, and `--write-runner-manifest` for generated write runners. `external` selects non-JS runners; `python` is kept as a compatibility alias. Write manifests use the same metadata shape as read manifests, but commands receive `--operation write --rows N --cols N --workload <name> --repeat N --warmup N --json`.

The generated `feature-rich` read workload uses the `read-values-rich-metadata` profile. Ascend opens this lane with `{ mode: 'values', richMetadata: true }`, keeping the pure `read-values` lane fast while requiring comments, hyperlinks, data validations, conditional formatting, and defined names for correctness.

Results vary by machine; record JSON or logs when publishing comparisons.

## Upstream benchmark profiles

Generated workloads are useful for coverage, but they are not enough for SOTA claims. `fixtures/benchmarks/upstream-profiles.ts` runs the competitive harness on dimensions published by other projects and annotates every case with `upstreamProfile`, `upstreamSourceLibrary`, `upstreamSourceBenchmark`, and `upstreamSourceUrl`.

For optimization loops, keep the first pass Ascend-only and phase-split before running a full external lane:

```bash
bun run bench:bottleneck:writer --rows 12800 --cols 50 --workload mixed-50pct-text --repeat 5 --warmup 1 --json
bun run bench:bottleneck:writer --rows 102400 --cols 50 --workload plain-text --repeat 3 --warmup 1 --string-mode inline --json
bun run bench:bottleneck:writer --rows 102400 --cols 50 --workload plain-text --repeat 3 --warmup 1 --string-mode plain --json
bun run bench:bottleneck:writer --profile excelize-generation-102400x50-plain-text --repeat 3 --warmup 1 --json
bun run bench:bottleneck:reader --profile fastexcel-reader-65536 --repeat 3 --warmup 1 --json
bun run bench:bottleneck:reader:fastexcel:read
bun run bench:bottleneck:reader:fastexcel:hydrate
bun run bench:zip-ab --phase parts --repeat 1 --fixture research/excel-corpus/NYC_311_SR_2010-2020-sample-1M.xlsx
```

The writer loop reports model-build, XLSX-write, optional validation, cells/sec, ns/cell, bytes, and memory separately. The reader loop reports direct ordered scan, `readXlsx`, materialization/hash validation, cells/sec, bytes, and memory separately. Use these phase loops to accept or revert candidate changes before spending time on the cross-language matrix. Use `--profile <upstream-profile>` to make the inner loop match a published external workload exactly. Use writer `--validate` only for confirmation samples because reopening large generated workbooks can dominate iteration time.

`bench:zip-ab --phase parts` is the tight loop for ZIP/XML read bottleneck work. It reports large XML part inflate, UTF-8 decode, string token scan, and byte token scan timing, and skips CRC/re-ZIP/write-side deflate tables so it stays usable on large real workbooks.

Dense streaming writer XML chunking can be A/B tested with `ASCEND_DENSE_XML_BYTE_BATCH` and `ASCEND_DENSE_XML_ROW_BATCH`. The default is tuned for the Excelize 102400 x 50 plain-text workload; larger chunks were slower and used more RSS on local Bun/JSC, while overly small chunks increased write overhead.

For external SOTA hillclimbs, use the profile-specific scripts as the inner loop and reserve the all-library sweeps for confirmation:

```bash
bun run bench:sota:write:excelize:fast
bun run bench:sota:write:pyopenxlsx
bun run bench:sota:read:fastexcel
bun run bench:sota:read:fastxlsx
bun run bench:sota:real:calamine:smoke
```

`bench:sota:write:excelize:fast` intentionally compares only Ascend and Excelize on Excelize's 102400 x 50 public workload. The full `bench:sota:write:excelize` script includes slower JS/Python/JVM baselines and can take several minutes inside a single external runner, so it is a confirmation sweep rather than a tight optimization loop.

Reader hillclimbs can run only the phase currently under optimization:

- `--phase read` times only `readXlsx` open/read on the pinned upstream shape.
- `--phase hydrate` times `readXlsx` plus materialization/hash correctness.
- `--phase direct` times the low-level ordered OOXML scan used for streaming validation.
- `--phase all` keeps the full phase split for confirmation.

Use `--gc-between-samples` for tighter memory feedback in short local loops, then rerun the external-process profile without that flag before publishing a scoreboard result.

Use `bench:profile:reader:fastexcel` and `bench:profile:writer:excelize` to run the aligned phase loops under Bun's profiler flags. Some Bun builds accept the profiling flags but emit no artifact; in that case `fixtures/benchmarks/profile-bun.ts` writes a `.profile-missing.txt` diagnostic and preserves the benchmark exit code. Add `--require-output` when the profile artifact itself is the gate.

```bash
bun run bench:upstream-profiles --profile-set write-smoke --repeat 5 --validation-mode final --json > upstream-write-smoke.json
bun run bench:upstream-profiles --profile-set read-smoke --repeat 5 --json > upstream-read-smoke.json
bun run bench:upstream-profiles --profile all --repeat 5 --json > upstream-profiles.json
bun run bench:competitive-scoreboard upstream-profiles.json --metric medianMs
```

`--profile-set write-smoke|read-smoke|write-heavy|docker-heavy` gives curated subsets so hillclimbs do not accidentally fan out into every runtime. Write profiles include the default cross-language writer set through `fixtures/benchmarks/runners/sota-writers.manifest.json`, currently Apache POI, Excelize, dhatim/fastexcel Java, ClosedXML, NPOI, Rust xlsxwriter, SheetJS, ExcelJS, XlsxWriter, PyExcelerate, and OpenPyXL; pass `--write-runner-manifest` to override that set. For heavyweight upstream shapes, prefer `--competitor external` with `--libraries` and a profile-specific manifest so the run still receives upstream annotations while avoiding unrelated lanes. When a library has both in-process and external-process adapters, add `--execution-scope external-process` to keep timing on the runner protocol path. Use `--validation-mode final` for hillclimb/candidate runs: samples still time the operation, and supported runners reopen/validate the final produced workbook once. Add `--timeout-ms <milliseconds>` for exploratory loops that include heavyweight external libraries; leave it off for final confirmation sweeps unless the timeout itself is part of the gate.

For fastexcel-reader-style comparisons, prefer the path-backed operation lane when comparing to fastexcel's published read shape:

```bash
bun run bench:upstream-profiles \
  --profile fastexcel-reader-65536 \
  --competitor external \
  --execution-scope external-process \
  --libraries ascend-readxlsx-raw-values-operation-path,fastexcel \
  --repeat 5 \
  --warmup 1 \
  --json > fastexcel-reader-profile.json
```

That lane times `readXlsx` open/read only and validates the final workbook after timing, matching fastexcel's runner timing model while preserving correctness checks. The ordered-hash lane remains useful for streaming validation stress tests, but it intentionally times hashing and is not the fair operation-timing comparison.

The path-backed operation lane uses `fixtures/benchmarks/runners/ascend_readxlsx_open_runner.ts`, which delegates each timed sample to a minimal `readXlsx` worker and performs ordered-hash correctness validation after timing. This keeps the RSS metric focused on the opened workbook rather than the benchmark adapter process. The default SparseGrid chunk size is 16x16; `ASCEND_CHUNK_BITS=5` or `6` can still be used for A/B checks on wider workloads.

```bash
bun run bench:upstream-profiles \
  --profile excelize-generation-102400x50-plain-text \
  --competitor external \
  --libraries ascend-external-writer,sheetjs,exceljs,xlsxwriter,xlsxwriter-constant-memory,pyexcelerate,pyexcelerate-range,pyexcelerate-cell,openpyxl,openpyxl-write-only,apache-poi,excelize \
  --write-runner-manifest /tmp/required-plain-text-writers.json \
  --repeat 5 \
  --warmup 1 \
  --validation-mode final \
  --json > excelize-profile.json
```

Profiles currently encoded:

| Profile | Source |
|---------|--------|
| `openpyxl-write-1000x50-10pct-text` | openpyxl write-performance docs: 1000 rows x 50 cols x 1 sheet, 10% text |
| `xlsxwriter-write-memory-{200,400,800,1600,3200,6400,12800}x50-50pct-text` | XlsxWriter memory/performance scaling table: N rows x 50 cols, 50/50 strings and numbers |
| `pyexcelerate-write-values-1000x100` | PyExcelerate value benchmark table, including bulk-sheet, range, and cell-by-cell runner lanes |
| `pyexcelerate-write-styles-1000x100` | PyExcelerate style benchmark table, including bulk-sheet, range, and cell-by-cell runner lanes |
| `apache-poi-ssperformance-xssf-50000x50` | Apache POI FAQ `SSPerformanceTest` XSSF generation check |
| `excelize-generation-102400x50-plain-text` | Excelize performance comparison: 102,400 x 50 plain text matrix generation |
| `fastexcel-writer-100000x4` | dhatim/fastexcel Java writer benchmark |
| `fastexcel-reader-65536` | fastexcel-reader 65,536-line read benchmark |

Primary sources:

- openpyxl performance docs: <https://openpyxl.readthedocs.io/en/3.0/performance.html>
- XlsxWriter memory/performance docs: <https://xlsxwriter.readthedocs.io/working_with_memory.html>
- PyExcelerate benchmark table: <https://github.com/kz26/PyExcelerate>
- Apache POI FAQ: <https://poi.apache.org/help/faq>
- Excelize performance docs: <https://xuri.me/excelize/en/performance.html>
- fastexcel README: <https://github.com/dhatim/fastexcel>

## Upstream real-workbook profiles

`fixtures/benchmarks/upstream-real-workbooks.ts` runs published real-workbook benchmarks through the competitive real-workbook harness. Unlike `upstream-profiles.ts`, it does not synthesize the workload during benchmark execution. The pinned source artifact must exist locally, and the harness rejects wrong byte sizes, wrong SHA-256 hashes, or child benchmark output that does not match the upstream shape contract.

The default runner manifest includes Ascend plus the installed cross-language SOTA reader set: OpenPyXL, fastexcel, python-calamine, rust-calamine, Excelize, Apache POI, ClosedXML, and Polars split by `calamine`, `xlsx2csv`, and `openpyxl` engines.

```bash
bun run bench:upstream-real --profile calamine-nyc311-1m --repeat 5 --warmup 3 --json > upstream-real.json
bun run bench:competitive-scoreboard upstream-real.json --metric medianMs
```

For a focused real-workbook SOTA loop against the calamine-published NYC 311 shape, use:

```bash
bun run bench:sota:real:calamine:smoke
bun run bench:sota:real:calamine
```

These use `fixtures/benchmarks/runners/nyc311-ordered-readers.manifest.json` to compare Ascend's ordered streaming values lane with rust-calamine on the same ordered-hash validation contract. The smoke variant runs one sample; the full script runs three samples.

For a fast external-runner smoke, use the pinned expected-shape sidecar and skip in-process JS adapters:

```bash
bun run bench:competitive-real --competitor external --category read --runner-manifest fixtures/benchmarks/runners/rust-calamine.manifest.json --expected-shape-sidecar fixtures/benchmarks/upstream-real-workbooks/nyc311-shape.json --repeat 1 --warmup 0 --json research/excel-corpus/NYC_311_SR_2010-2020-sample-1M.xlsx
```

Current profile:

| Profile | Source | Local path |
|---------|--------|------------|
| `calamine-nyc311-1m` | calamine README performance benchmark: one worksheet, 1,000,001 rows x 41 columns, 28,056,975 non-empty cells | `research/excel-corpus/NYC_311_SR_2010-2020-sample-1M.xlsx` |

Source workbook acquisition:

- Dataset archive: <https://raw.githubusercontent.com/wiki/jqnatividad/qsv/files/NYC_311_SR_2010-2020-sample-1M.7z>
- Published benchmark definition: <https://docs.rs/crate/calamine/latest/source/README.md>
- Archive pin: `48111517` bytes, SHA-256 `5c5f876b097ed6b51d52a5309c029ac605e959204cfb64a41f847bdc3ef3165b`
- Extracted CSV pin: `538951068` bytes, SHA-256 `18f0dd774a6c4b79da3dbf3aa0cd878d374dab132226af2c629d9eef9595061b`
- Materialize XLSX:
  `python3 fixtures/benchmarks/materialize_nyc311_xlsx.py <csv> research/excel-corpus/NYC_311_SR_2010-2020-sample-1M.xlsx --json`
- Materialized XLSX pin: `249316631` bytes, SHA-256 `74a9b50621cf9b0fe8cdb2d4072b5535a2c0e2d83247bb38a37a3b3d809202ea`
- Expected-shape sidecar: `fixtures/benchmarks/upstream-real-workbooks/nyc311-shape.json`

The calamine README reports an XLSX filesize of `186MB`, while the public download is a compressed CSV. Ascend pins the deterministic XlsxWriter 3.2.9 materialization separately so benchmark results do not overstate byte-identical reproduction of the original calamine machine-local XLSX.

Required local runtimes for the full default manifest include Python packages from the runner install hints, Rust/Cargo, Java/Maven, .NET 8, and Go (`brew install go` for Excelize).

For cleaner repeatability, the core SOTA runner environment can be built from `fixtures/benchmarks/runners/Dockerfile.sota`. It includes Bun, Python writer/read libraries, Rust/Cargo, Go, Java/Maven, .NET 8, `hyperfine`, and archive tooling. The wrapper keeps dependency caches in named Docker volumes while bind-mounting the working tree:

```bash
bun run bench:sota:docker:build
bun run bench:sota:docker:bootstrap
bun run bench:sota:docker -- bun run bench:upstream-profiles --profile excelize-generation-102400x50-plain-text --competitor external --validation-mode final --repeat 5 --json
```

The bootstrap step runs `bun install --frozen-lockfile`, verifies pinned Python runner packages, fetches/builds Rust and Go runners, warms Maven, restores ClosedXML, and prints runtime versions. Docker keeps dependency/build caches in named volumes and `.dockerignore` keeps local benchmark artifacts and dependency directories out of the image build context.

For more repeatable local measurements, set `ASCEND_SOTA_CPUSET` and/or `ASCEND_SOTA_MEMORY` before invoking the wrapper; these pass through to Docker as CPU and memory limits.

NPOI remains license-gated: installing .NET is fine, but do not enable or publish NPOI runs unless `ACCEPT_NPOI_OSMF_LICENSE=1` has been set after explicit review.

## Formula correctness vs HyperFormula

```bash
bun run fixtures/formulas/formula-hyperformula-compare.ts
```

Reports mismatches on a small shared scenario set (informational; engines differ on edge cases).

## Excel ground truth (manual)

```bash
bun run excel:ground-truth
```

Prints instructions for exporting expected values from Excel and comparing to JSON fixtures under `fixtures/formulas/`.
