# Performance Claim Baseline Matrix

Updated: 2026-05-15

## Current Claim Gate

Status: defer.

This document is the tracked release claim artifact. `/private/tmp` paths below
are reproducibility pointers to the clean run outputs, not required context for
release wording.

No broad XLSX read, SOTA, or QSS-leapfrog speed claim is promotable from this artifact. The run below is useful external baseline evidence for one release workflow, but it is not a clean full-profile claim because:

- `competitive-scoreboard --require-profile xlsx-read-sota` fails coverage for the rest of the profile.
- The run covered one public/reproducible generated workload: `string-heavy`, `read-values`, `raw-ooxml`.
- Several external runners were unavailable or blocked in the clean benchmark worktree. They are recorded as blockers, not wins.
- Several timing lanes are semantically related but not one unified timing boundary. Do not collapse in-process, preloaded-bytes, file-path, row-stream, and materialized-workbook timings into a single "wins everything" claim.

Humble allowed wording:

> On the generated `string-heavy` 2000x20 raw OOXML workload, using `fixtures/benchmarks/competitive-io.ts` with 5 samples and 1 warmup from a clean detached worktree, Ascend's value-read paths were competitive with or faster than the external readers that successfully ran. Several external runners were unavailable or blocked, so this is one scoped baseline row, not a full `xlsx-read-sota` claim.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend is SOTA for XLSX read."
- "Ascend beats every external library."
- Any wording that treats failed or unavailable runners as wins.

Next action: defer the speed claim. Expand this into a clean-tree `xlsx-read-sota` matrix across the required workloads and unblock missing runner dependencies, then optimize only if a named release workflow loses or shows an unstable tail.

## Owner-Ready Benchmark Blocker

Owner: benchmarking/external baselines.

Blocker: broad read-speed and QSS-leapfrog performance wording is blocked until
the full `xlsx-read-sota` profile is reproduced from a clean worktree or
downgraded per runner/workload with explicit blocker reasons.

Exact next command shape:

```bash
bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload <xlsx-read-sota-workload> --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json
bun run fixtures/benchmarks/competitive-scoreboard.ts <suite.json> --json --metric medianMs --require-profile xlsx-read-sota
```

Acceptance evidence:

- Clean detached worktree or clean release benchmark environment.
- Every `xlsx-read-sota` workload either has comparable Ascend and external rows
  or an explicit `runner unavailable`, `blocked`, `unsupported-operation`, or
  `not comparable` status.
- Median, p95, CV/noise, memory, environment, runner/library versions, command,
  input shape, and semantic comparability are recorded for each comparable row.
- Failed, missing, or semantically mismatched runners are not counted as wins.

Stop condition: do not optimize from this single `string-heavy` row. Optimize
only after the full profile identifies one release workflow as a meaningful
loss, unstable tail, or memory/latency tradeoff worth production work.

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

Optimize: no production optimization from this single row. The top release action is not another micro-optimization; it is clean, broader external evidence.

Kill: no current optimization is killed by this row.

Defer: yes. Defer all broad read-speed wording until a clean-tree full `xlsx-read-sota` run either passes coverage or records per-runner blockers without counting them as wins. The immediate next action is runner-environment hardening for `fastexcel`, `python-calamine`, Polars engines, `pyopenxlsx`, `fastxlsx`, and ClosedXML, followed by the full profile run.
