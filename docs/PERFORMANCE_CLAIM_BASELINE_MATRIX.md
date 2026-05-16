# Performance Claim Baseline Matrix

Updated: 2026-05-15

## Current Claim Gate

Status: defer.

This document is the tracked release claim artifact. `/private/tmp` paths below
are reproducibility pointers to the clean run outputs, not required context for
release wording.

No broad XLSX read, SOTA, or QSS-leapfrog speed claim is promotable from this artifact. The cycles below are useful external baseline evidence for scoped release workflows, but they are not a clean full-profile claim because:

- `competitive-scoreboard --require-profile xlsx-read-sota` fails coverage for the rest of the profile.
- The recorded cycles cover public/reproducible generated `dense-values`, `sparse-wide`, `styles-heavy`, and `string-heavy` `read-values` workloads over `raw-ooxml`, not the full profile.
- Several external runners were unavailable or blocked in the clean benchmark worktree. They are recorded as blockers, not wins.
- Several timing lanes are semantically related but not one unified timing boundary. Do not collapse in-process, preloaded-bytes, file-path, row-stream, and materialized-workbook timings into a single "wins everything" claim.

Humble allowed wording:

> On the generated `string-heavy` 2000x20 raw OOXML workload, using `fixtures/benchmarks/competitive-io.ts` with 5 samples and 1 warmup from a clean detached worktree, Ascend's value-read paths were competitive with or faster than the external readers that successfully ran. Several external runners were unavailable or blocked, so this is one scoped baseline row, not a full `xlsx-read-sota` claim.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend is SOTA for XLSX read."
- "Ascend beats every external library."
- Any wording that treats failed or unavailable runners as wins.

Next action: defer the speed claim. Continue expanding the clean-tree `xlsx-read-sota` matrix across the required workloads and unblock remaining runner issues, then optimize only if a named release workflow loses or shows an unstable tail.

## Owner-Ready Benchmark Blocker

Owner: benchmarking/external baselines.

Blocker: broad read-speed and QSS-leapfrog performance wording is blocked until
the full `xlsx-read-sota` profile is reproduced from a clean worktree or
downgraded per runner/workload with explicit blocker reasons.

Exact next command shape:

```bash
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload <xlsx-read-sota-workload> --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts <suite.json> --json --metric medianMs --require-profile xlsx-read-sota
```

Acceptance evidence:

- Clean detached worktree or clean release benchmark environment.
- Every `xlsx-read-sota` workload either has comparable Ascend and external rows
  or an explicit `runner unavailable`, `blocked`, `unsupported-operation`, or
  `not comparable` status.
- Median, p95, CV/noise, memory, environment, runner/library versions, command,
  input shape, and semantic comparability are recorded for each comparable row.
- Failed, missing, or semantically mismatched runners are not counted as wins.

Stop condition: do not optimize from the partial `dense-values`, `sparse-wide`,
`styles-heavy`, and `string-heavy` rows. Optimize only after the full profile
identifies one release workflow as a meaningful loss, unstable tail, or
memory/latency tradeoff worth production work.

## Cycle: Dense Value Read

Classification: defer. No production optimization is justified from this cycle.

Workflow: XLSX open/inspect value read for a dense numeric worksheet.

Why it matters for release: this is a required `xlsx-read-sota` workload and a common agent inspection path for tabular numeric workbooks.

Public/tracked-clean input: `competitive-io` generated `dense-values` workbook using tracked benchmark code from detached commit `136c1324`, `raw-ooxml` source, 2000 rows x 20 columns, 40,000 logical cells, 112,699 input bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-136c1324 136c1324
cd /private/tmp/ascend-perf-hillclimb-136c1324
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload dense-values --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-136c1324-runs/dense-values-read-values-v2.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-136c1324-runs/dense-values-read-values-v2.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-136c1324-runs/dense-values-scoreboard-v2.json
```

Environment:

- Commit: `136c1324`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-136c1324`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Python: `3.13.3`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-136c1324-runs/dense-values-read-values-v2.json
/private/tmp/ascend-perf-hillclimb-136c1324-runs/dense-values-scoreboard-v2.json
```

All successful timing rows below use 5 measured samples after 1 warmup. Rows marked `blocked` or `runner unavailable` are non-ranking status rows from the harness and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 5.757 | 14.262 | 0.517 | 85.6 MiB | File-path value read with post-operation assertions. Comparable for value-read inspection only; tail is noisy. |
| Ascend row stream | not comparable | `ascend-readxlsx-row-stream-bytes` | 8.817 | 10.418 | 0.108 | 86.4 MiB | Row-stream over bytes; useful diagnostic path but not the default file-path comparison. |
| fastexcel Python | ran/lost | `fastexcel` | 8.860 | 16.373 | 0.332 | 63.8 MiB | Value read, operation timing. Lower memory; bounded gap versus Ascend file-path is 1.54x slower by median. |
| Polars calamine | ran/lost | `polars-calamine` | 9.361 | 10.268 | 0.060 | 103.6 MiB | Value read through Polars/calamine. Lower tail than Ascend in this run; bounded gap is 1.63x slower by median. |
| FastExcel Java | ran/lost | `fastexcel-java` | 14.884 | 33.707 | 0.504 | 96.0 MiB heap | Streaming value read. Tail is noisy; bounded gap is 2.59x slower by median. |
| python-calamine | ran/lost | `python-calamine` | 18.987 | 22.590 | 0.088 | 41.7 MiB | File-path materialized value read. Lower memory than Ascend; 3.30x slower by median. |
| rust-calamine | ran/lost | `rust-calamine` | 38.792 | 41.892 | 0.040 | 13.9 MiB | File-path materialized value read. Much lower memory; 6.74x slower by median. |
| SheetJS | ran/lost | `sheetjs` | 41.323 | 47.644 | 0.070 | 254.5 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| ExcelJS | ran/lost | `exceljs` | 68.534 | 136.152 | 0.419 | 294.6 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| Polars xlsx2csv | ran/lost | `polars-xlsx2csv` | 74.998 | 75.661 | 0.013 | 75.7 MiB | Value read through xlsx2csv path. 13.03x slower by median. |
| Apache POI | ran/lost | `apache-poi` | 92.890 | 157.859 | 0.352 | 232.0 MiB heap | Materialized workbook value read. Tail is noisy; 16.14x slower by median. |
| Excelize | ran/lost | `excelize` | 119.834 | 123.387 | 0.016 | 42.3 MiB | File-path materialized value read. Lower memory; 20.81x slower by median. |
| Polars openpyxl | ran/lost | `polars-openpyxl` | 148.253 | 162.503 | 0.093 | 119.0 MiB | Value read through openpyxl engine. 25.75x slower by median. |
| openpyxl read-only values | ran/lost | `openpyxl-read-only-values` | 247.894 | 333.696 | 0.220 | 64.6 MiB | Streaming data-only value read. Lower memory; 43.06x slower by median. |
| pyopenxlsx | ran/lost | `pyopenxlsx` | 347.192 | 548.419 | 0.397 | 74.6 MiB | Cell materialization. Tail is noisy; 60.31x slower by median. |
| ClosedXML | blocked | `closedxml` | n/a | n/a | n/a | n/a | Not counted. The .NET runner failed with `CSSM_ModuleLoad(): One or more parameters passed to a function were not valid.` |
| fastxlsx | runner unavailable | `fastxlsx` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `fastxlsx` is not installed in the clean Python environment. |

Coverage gate result: failed, as expected for a partial profile. The dense row still lacks ClosedXML coverage because the runner is blocked. The profile remains missing `sparse-wide`, `string-heavy` from current commit, `styles-heavy`, `formula-heavy`, `table-heavy`, `feature-rich`, `selected-sheet`, `metadata-only`, and `warm-workflow` coverage.

Humble allowed wording:

> On the generated `dense-values` 2000x20 raw OOXML workload at commit `136c1324`, Ascend's file-path value-read row had the fastest eligible median among runners that completed, but its tail was noisy and the full `xlsx-read-sota` profile is still incomplete.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend has a clean dense-read speed win across all runners."
- "Ascend beats ClosedXML or fastxlsx" from this run.
- Any wording that ignores the noisy Ascend p95/CV or the incomplete full profile.

Next action: defer production optimization. Continue profile expansion with `sparse-wide` using the same explicit PATH, and keep ClosedXML/fastxlsx as blockers unless their runners are fixed.

## Cycle: Sparse-Wide Value Read

Classification: defer. No production optimization is justified from this cycle.

Workflow: XLSX open/inspect value read for a sparse, wide worksheet.

Why it matters for release: this is a required `xlsx-read-sota` workload and a common shape for operational spreadsheets with many possible columns but sparse populated cells.

Public/tracked-clean input: `competitive-io` generated `sparse-wide` workbook using tracked benchmark code from detached commit `2e71900f`, `raw-ooxml` source, 5000 rows x 256 columns, 1,280,000 logical cells, 23,093 populated cells, 166,216 input bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-2e71900f 2e71900f
cd /private/tmp/ascend-perf-hillclimb-2e71900f
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload sparse-wide --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-2e71900f-runs/sparse-wide-read-values.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-2e71900f-runs/sparse-wide-read-values.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-2e71900f-runs/sparse-wide-scoreboard.json
```

Environment:

- Commit: `2e71900f`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-2e71900f`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Python: `3.13.3`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-2e71900f-runs/sparse-wide-read-values.json
/private/tmp/ascend-perf-hillclimb-2e71900f-runs/sparse-wide-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. Rows marked `blocked` or `runner unavailable` are non-ranking status rows from the harness and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 11.651 | 20.132 | 0.344 | 120.9 MiB | File-path value read with post-operation assertions. Comparable for value-read inspection only; tail is moderately noisy. |
| Ascend row stream | not comparable | `ascend-readxlsx-row-stream-bytes` | 9.506 | 12.146 | 0.160 | 100.5 MiB | Row-stream over bytes; useful diagnostic path but not the default file-path comparison. |
| SheetJS | ran/lost | `sheetjs` | 23.217 | 25.338 | 0.042 | 348.5 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| fastexcel Python | ran/lost | `fastexcel` | 31.264 | 34.354 | 0.058 | 125.7 MiB | Value read, operation timing. Bounded gap versus Ascend file-path is 2.68x slower by median. |
| python-calamine | ran/lost | `python-calamine` | 37.319 | 47.433 | 0.133 | 95.0 MiB | File-path materialized value read. Lower memory; 3.20x slower by median. |
| ExcelJS | ran/lost | `exceljs` | 37.645 | 47.602 | 0.122 | 391.6 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| rust-calamine | ran/lost | `rust-calamine` | 79.400 | 130.920 | 0.362 | 46.7 MiB | File-path materialized value read. Much lower memory; tail is noisy; 6.81x slower by median. |
| Polars calamine | ran/lost | `polars-calamine` | 100.916 | 136.013 | 0.325 | 205.6 MiB | Value read through Polars/calamine. Tail is noisy; 8.66x slower by median. |
| Polars xlsx2csv | ran/lost | `polars-xlsx2csv` | 211.594 | 248.661 | 0.081 | 217.7 MiB | Value read through xlsx2csv path. 18.16x slower by median. |
| openpyxl read-only values | ran/lost | `openpyxl-read-only-values` | 235.459 | 247.074 | 0.026 | 58.3 MiB | Streaming data-only value read. Lower memory; 20.21x slower by median. |
| Apache POI | ran/lost | `apache-poi` | 299.826 | 582.300 | 0.436 | 344.0 MiB heap | Materialized workbook value read. Tail is noisy; 25.73x slower by median. |
| Excelize | ran/lost | `excelize` | 1048.502 | 1060.777 | 0.011 | 635.8 MiB | File-path materialized value read. 90.00x slower by median. |
| Polars openpyxl | ran/lost | `polars-openpyxl` | 1574.377 | 2194.467 | 0.185 | 486.1 MiB | Value read through openpyxl engine. Tail is noisy; 135.13x slower by median. |
| pyopenxlsx | ran/lost | `pyopenxlsx` | 20405.904 | 31269.832 | 0.250 | 229.4 MiB | Cell materialization. Extremely slow on this sparse-wide shape; 1751.42x slower by median. |
| FastExcel Java | blocked | `fastexcel-java` | n/a | n/a | n/a | n/a | Not counted. Runner exited with code 1 on this workload. |
| ClosedXML | blocked | `closedxml` | n/a | n/a | n/a | n/a | Not counted. The .NET runner failed with `CSSM_ModuleLoad(): One or more parameters passed to a function were not valid.` |
| fastxlsx | runner unavailable | `fastxlsx` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `fastxlsx` is not installed in the clean Python environment. |

Coverage gate result: failed, as expected for a partial profile. The sparse-wide row lacks FastExcel Java, ClosedXML, and fastxlsx coverage. The profile remains missing `string-heavy` from current commit, `styles-heavy`, `formula-heavy`, `table-heavy`, `feature-rich`, `selected-sheet`, `metadata-only`, and `warm-workflow` coverage.

Humble allowed wording:

> On the generated `sparse-wide` 5000x256 raw OOXML workload at commit `2e71900f`, Ascend's file-path value-read row had the fastest eligible median among completed comparable external runners. The full `xlsx-read-sota` profile remains incomplete, and several runner blockers are still recorded.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend beats FastExcel Java, ClosedXML, or fastxlsx" from this run.
- Any wording that hides the very slow but valid pyopenxlsx row, the blocked runners, or the incomplete full profile.

Next action: defer production optimization. Continue profile expansion with `styles-heavy`; separately investigate the FastExcel Java sparse-wide failure and ClosedXML runtime failure as runner blockers.

## Cycle: Styles-Heavy Value Read

Classification: defer. No production optimization is justified from this cycle.

Workflow: XLSX open/inspect value read for a style-dense worksheet.

Why it matters for release: this is a required `xlsx-read-sota` workload and a common shape for business workbooks where formatting is widespread even when the release workflow only needs fast value inspection.

Public/tracked-clean input: `competitive-io` generated `styles-heavy` workbook using tracked benchmark code from detached commit `b1a53ea0`, `raw-ooxml` source, 2000 rows x 20 columns, 40,000 logical cells, 40,000 populated cells, 257,177 input bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-b1a53ea0 b1a53ea0
cd /private/tmp/ascend-perf-hillclimb-b1a53ea0
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload styles-heavy --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-b1a53ea0-runs/styles-heavy-read-values.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-b1a53ea0-runs/styles-heavy-read-values.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-b1a53ea0-runs/styles-heavy-scoreboard.json
```

Environment:

- Commit: `b1a53ea0`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-b1a53ea0`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Python: `3.13.3`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-b1a53ea0-runs/styles-heavy-read-values.json
/private/tmp/ascend-perf-hillclimb-b1a53ea0-runs/styles-heavy-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. Rows marked `blocked`, `runner unavailable`, or `not comparable` are non-ranking status rows for claim wording and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 5.733 | 7.387 | 0.142 | 99.4 MiB | File-path value read with post-operation assertions. Comparable for value-read inspection only; this does not prove style fidelity or preservation. |
| Ascend row stream | not comparable | `ascend-readxlsx-row-stream-bytes` | 6.012 | 6.161 | 0.023 | 107.8 MiB | Row-stream over bytes; useful diagnostic path but not the default file-path comparison. |
| fastexcel Python | ran/lost | `fastexcel` | 9.518 | 9.581 | 0.017 | 63.3 MiB | Value read, operation timing. Lower memory; bounded gap versus Ascend file-path is 1.66x slower by median. |
| Polars calamine | ran/lost | `polars-calamine` | 11.099 | 12.675 | 0.069 | 104.0 MiB | Value read through Polars/calamine. Bounded gap is 1.94x slower by median. |
| FastExcel Java | ran/lost | `fastexcel-java` | 12.527 | 28.055 | 0.427 | 112.0 MiB heap | Streaming value read. Tail is noisy; bounded gap is 2.18x slower by median. |
| python-calamine | ran/lost | `python-calamine` | 19.725 | 22.513 | 0.064 | 42.2 MiB | File-path materialized value read. Lower memory; 3.44x slower by median. |
| rust-calamine | ran/lost | `rust-calamine` | 38.072 | 39.113 | 0.017 | 11.5 MiB | File-path materialized value read. Much lower memory; 6.64x slower by median. |
| SheetJS | ran/lost | `sheetjs` | 39.359 | 40.306 | 0.026 | 301.7 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| ExcelJS | ran/lost | `exceljs` | 53.636 | 59.321 | 0.054 | 311.5 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| Apache POI | ran/lost | `apache-poi` | 73.070 | 193.725 | 0.574 | 352.0 MiB heap | Materialized workbook value read. Tail is noisy; 12.75x slower by median. |
| Polars xlsx2csv | ran/lost | `polars-xlsx2csv` | 78.880 | 80.045 | 0.009 | 75.7 MiB | Value read through xlsx2csv path. 13.76x slower by median. |
| Excelize | ran/lost | `excelize` | 135.631 | 163.217 | 0.096 | 47.4 MiB | File-path materialized value read. Lower memory; 23.66x slower by median. |
| Polars openpyxl | ran/lost | `polars-openpyxl` | 148.769 | 163.828 | 0.079 | 120.1 MiB | Value read through openpyxl engine. 25.95x slower by median. |
| openpyxl read-only values | ran/lost | `openpyxl-read-only-values` | 167.175 | 171.429 | 0.014 | 64.3 MiB | Streaming data-only value read. Lower memory; 29.16x slower by median. |
| pyopenxlsx | ran/lost | `pyopenxlsx` | 229.439 | 233.102 | 0.010 | 75.7 MiB | Cell materialization. 40.02x slower by median. |
| ClosedXML | blocked | `closedxml` | n/a | n/a | n/a | n/a | Not counted. The .NET runner failed with `CSSM_ModuleLoad(): One or more parameters passed to a function were not valid.` |
| fastxlsx | runner unavailable | `fastxlsx` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `fastxlsx` is not installed in the clean Python environment. |

In-process JavaScript reference:

| Library | Status | Median ms | P95 ms | CV | Peak RSS |
| --- | --- | ---: | ---: | ---: | ---: |
| Ascend in-process | ran/won | 4.727 | 4.947 | 0.029 | 199.8 MiB |
| SheetJS `xlsx@0.18.5` | ran/lost | 39.359 | 40.306 | 0.026 | 301.7 MiB |
| ExcelJS `4.4.0` | ran/lost | 53.636 | 59.321 | 0.054 | 311.5 MiB |

Coverage gate result: failed, as expected for a partial profile. The styles-heavy row lacks ClosedXML and fastxlsx coverage. The profile remains missing `string-heavy` from current commit, `formula-heavy`, `table-heavy`, `feature-rich`, `selected-sheet`, `metadata-only`, and `warm-workflow` coverage.

Humble allowed wording:

> On the generated `styles-heavy` 2000x20 raw OOXML workload at commit `b1a53ea0`, Ascend's file-path value-read row had the fastest eligible median and p95 among completed comparable external runners. The full `xlsx-read-sota` profile remains incomplete, and ClosedXML/fastxlsx blockers are still recorded.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend has a clean style-heavy read speed win across all runners."
- "Ascend beats ClosedXML or fastxlsx" from this run.
- "Ascend proves style fidelity or preservation" from this value-read run.
- Any wording that hides the incomplete full profile or blocked/unavailable runners.

Next action: defer production optimization. Continue profile expansion with `formula-heavy`; keep ClosedXML and fastxlsx as blockers unless their runners are fixed.

## Workflow

Workflow: XLSX open/inspect value read for a string-heavy worksheet.

Why it matters for release: this is the first user-visible agent workflow after receiving an unknown workbook: open the workbook, inspect cell values, and make a plan. It is also the broadest external baseline surface because SheetJS, ExcelJS, openpyxl, Calamine, Apache POI, ClosedXML, Excelize, Polars, and related runners can reasonably perform value reads.

Claim profile: partial `xlsx-read-sota`.

Operation profile: `read-values`.

Input: `competitive-io` generated `string-heavy` workbook using tracked benchmark code from detached commit `7ebaf8bb`, `raw-ooxml` source, 2000 rows x 20 columns, 40,000 logical cells, 215,952 input bytes. No private corpus or local research workbook was used. The generated temp XLSX is reproducible from the command below rather than checked in.

Semantic boundary: value-read correctness only. This does not prove formula calculation, style fidelity, package preservation, macro/signature handling, edit safety, save/reopen safety, or full workbook feature compatibility.

## Commands

Benchmark command:

```bash
git worktree add --detach /private/tmp/ascend-claim-clean-7ebaf8bb 7ebaf8bb
cd /private/tmp/ascend-claim-clean-7ebaf8bb
bun install --frozen-lockfile
bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload string-heavy --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-external-claim-matrix-2026-05-15-clean/string-heavy-read-values-clean.json
```

Coverage gate command:

```bash
bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-external-claim-matrix-2026-05-15-clean/string-heavy-read-values-clean.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-external-claim-matrix-2026-05-15-clean/string-heavy-read-values-scoreboard.json
```

Coverage gate result: failed, as expected for a single-workload row. The gate reported missing `xlsx-read-sota` coverage for `dense-values`, `sparse-wide`, `styles-heavy`, `formula-heavy`, `table-heavy`, `feature-rich`, `selected-sheet`, `metadata-only`, and `warm-workflow`. It also reported missing ClosedXML coverage for `string-heavy` because that runner failed in the clean worktree.

Raw output path:

```text
/private/tmp/ascend-external-claim-matrix-2026-05-15-clean/string-heavy-read-values-clean.json
/private/tmp/ascend-external-claim-matrix-2026-05-15-clean/string-heavy-read-values-scoreboard.json
```

## Environment

- Commit: `7ebaf8bb`
- Worktree: clean detached worktree at `/private/tmp/ascend-claim-clean-7ebaf8bb`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.6`
- Python: `3.9.6`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`
- JavaScript library versions: SheetJS `xlsx@0.18.5`, ExcelJS `4.4.0`

## Baseline Matrix

Reference comparison for external value readers: `ascend-readxlsx-raw-values-operation-path`, file-path input, `read-values`, median 9.347 ms, p95 16.545 ms, CV 0.300, peak RSS 129.8 MiB. This is the most directly reusable external-process Ascend path row from the run, but the tail is noisy and must not be over-promoted. `ascend-readxlsx-row-stream-bytes` was faster at median 8.609 ms, but its timing boundary is row-stream over bytes, so do not use it as the default cross-library headline.

All successful timing rows below use 5 measured samples after 1 warmup. Rows marked `runner unavailable` or `blocked` are non-ranking status rows from the harness and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 9.347 | 16.545 | 0.300 | 129.8 MiB | File-path value read with post-operation assertions. Comparable for value-read inspection only; tail is noisy. |
| Ascend row stream | ran/won | `ascend-readxlsx-row-stream-bytes` | 8.609 | 8.808 | 0.029 | 130.3 MiB | Faster bytes/row-stream path. Useful workflow evidence, not the default cross-library file-path comparison. |
| FastExcel Java | ran/lost | `fastexcel-java` | 12.694 | 27.138 | 0.397 | 344.0 MiB heap | Streaming value read. Tail is noisy; bounded gap versus Ascend file-path is 1.36x slower by median. |
| rust-calamine | ran/lost | `rust-calamine` | 45.367 | 45.941 | 0.009 | 14.7 MiB | File-path materialized value read. Much lower memory; 4.85x slower by median. |
| SheetJS | ran/lost | `sheetjs` | 58.637 | 60.634 | 0.022 | 426.0 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| ExcelJS | ran/lost | `exceljs` | 61.550 | 66.950 | 0.061 | 482.7 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| Apache POI | ran/lost | `apache-poi` | 90.558 | 216.948 | 0.550 | 352.0 MiB heap | Materialized workbook value read. Tail is noisy; 9.69x slower by median. |
| Excelize | ran/lost | `excelize` | 200.211 | 217.773 | 0.057 | 69.4 MiB | File-path materialized value read. 21.42x slower by median. |
| openpyxl read-only values | ran/lost | `openpyxl-read-only-values` | 297.690 | 312.419 | 0.034 | 59.3 MiB | Streaming data-only value read. Lower memory than Ascend; 31.85x slower by median. |
| fastexcel Python | runner unavailable | `fastexcel` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `fastexcel` is not installed in the clean Python environment. |
| python-calamine | runner unavailable | `python-calamine` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `python_calamine` is not installed in the clean Python environment. |
| Polars calamine | runner unavailable | `polars-calamine` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `polars` is not installed in the clean Python environment. |
| Polars xlsx2csv | runner unavailable | `polars-xlsx2csv` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `polars` is not installed in the clean Python environment. |
| Polars openpyxl | runner unavailable | `polars-openpyxl` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `polars` is not installed in the clean Python environment. |
| pyopenxlsx | runner unavailable | `pyopenxlsx` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `pyopenxlsx` is not installed in the clean Python environment. |
| fastxlsx | runner unavailable | `fastxlsx` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `fastxlsx` is not installed in the clean Python environment. |
| ClosedXML | blocked | `closedxml` | n/a | n/a | n/a | n/a | Not counted. The .NET runner failed with `CSSM_ModuleLoad(): One or more parameters passed to a function were not valid.` |

In-process JavaScript reference:

| Library | Status | Median ms | P95 ms | CV | Peak RSS |
| --- | --- | ---: | ---: | ---: | ---: |
| Ascend in-process | ran/won | 9.494 | 11.355 | 0.112 | 308.8 MiB |
| SheetJS `xlsx@0.18.5` | ran/lost | 58.637 | 60.634 | 0.022 | 426.0 MiB |
| ExcelJS `4.4.0` | ran/lost | 61.550 | 66.950 | 0.061 | 482.7 MiB |

Diagnostic note: an earlier run from the dirty main worktree had more Python/.NET runners available and is intentionally not used for claim wording. It can help prioritize runner setup, but not speed promotion.

## Decision

Promote: no.

Optimize: no production optimization from the partial profile rows. The top release action is not another micro-optimization; it is clean, broader external evidence and runner hardening.

Kill: no current optimization is killed by this row.

Defer: yes. Defer all broad read-speed wording until a clean-tree full `xlsx-read-sota` run either passes coverage or records per-runner blockers without counting them as wins. The immediate next action is `formula-heavy` profile expansion plus runner hardening for FastExcel Java on `sparse-wide`, ClosedXML, and fastxlsx, followed by the full profile run.
