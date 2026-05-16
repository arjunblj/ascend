# Performance Claim Baseline Matrix

Updated: 2026-05-16

## Current Claim Gate

Status: defer.

This document is the tracked release claim artifact. `/private/tmp` paths below
are reproducibility pointers to the clean run outputs, not required context for
release wording.

No broad XLSX read, SOTA, or QSS-leapfrog speed claim is promotable from this artifact. The cycles below are useful external baseline evidence for scoped release workflows, but they are not a clean full-profile claim because:

- `competitive-scoreboard --require-profile xlsx-read-sota` still fails for broad promotion. The current full-profile run at `9ddfff91` reports no leader failures, and a merged selected-sheet/metadata-only scoreboard removes those same-lane comparability failures, but ClosedXML coverage, feature-rich semantic mismatches, and unsupported selected-sheet/metadata-only competitors remain explicit blockers.
- The recorded cycles cover public/reproducible generated `dense-values`, `sparse-wide`, `styles-heavy`, `formula-heavy`, `table-heavy`, `feature-rich`, `selected-sheet`, `metadata-only`, `warm-workflow`, and `string-heavy` workloads over `raw-ooxml`, but they are per-workload evidence rows rather than one clean all-workload promotion run.
- Current harness evidence now supports same-lane selected-sheet rows for Ascend, SheetJS, OpenPyXL, and python-calamine. Treat older `openpyxl` and Calamine selected-sheet `unsupported-operation` wording as historical for the recorded clean runs.
- Current harness evidence now supports same-lane metadata-only rows for Ascend, SheetJS, OpenPyXL, and python-calamine. Calamine wins that head-to-head; treat older metadata-only `missing-comparable` or Calamine `unsupported-operation` wording as historical.
- Current harness evidence now supports a SheetJS feature-rich rich-metadata row using SheetJS `bookFiles`; older SheetJS `semantic-mismatch` wording is historical for the pre-runner-fix cycles. Calamine-family rich-metadata rows remain not comparable.
- Several external runners were unavailable or blocked in the clean benchmark worktree. They are recorded as blockers, not wins.
- Several timing lanes are semantically related but not one unified timing boundary. Do not collapse in-process, preloaded-bytes, file-path, row-stream, and materialized-workbook timings into a single "wins everything" claim.

Humble allowed wording:

> On the generated `string-heavy` 2000x20 raw OOXML workload, using `fixtures/benchmarks/competitive-io.ts` with 5 samples and 1 warmup from a clean detached worktree, Ascend's value-read paths were competitive with or faster than the external readers that successfully ran. Several external runners were unavailable or blocked, so this is one scoped baseline row, not a full `xlsx-read-sota` claim.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend is SOTA for XLSX read."
- "Ascend beats every external library."
- Any wording that treats failed or unavailable runners as wins.

Next action: downgrade the broad speed claim and stop production optimization from winning rows. Continue only if the performance loop is explicitly attacking a remaining claim blocker or measured loss: ClosedXML coverage, feature-rich semantic mismatches for SheetJS/Calamine, metadata-only versus Calamine, remaining unsupported selected-sheet/metadata-only competitors, or FastXLSX environment coverage.

## Metadata-Only Calamine Boundary

Classification: defer/optimization target. Calamine is now a comparable
metadata-only runner, and Ascend loses this row. A narrow attempted optimization
to skip plain worksheet relationship probes was rejected because paired local
measurement did not validate an improvement.

Workflow: generated XLSX metadata-only open/inspect, loading workbook and sheet
metadata without hydrating cells.

Why it matters for release: metadata-only open is the safety-first path used
before edit planning on unknown workbooks. A speed claim here must not ignore the
native Calamine baseline.

Public/tracked-clean input: `competitive-io` generated `metadata-only`
`raw-ooxml`, 200 rows x 20 columns, three workbook sheets, 15,347 input bytes,
from tracked benchmark code at commit `b6925afe`.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-metadata-calamine-clean-b6925afe b6925afe
cd /private/tmp/ascend-metadata-calamine-clean-b6925afe
bun install --frozen-lockfile
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only,sheetjs-metadata-only,openpyxl-metadata-only,python-calamine-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 15 --warmup 3 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-calamine-clean-b6925afe/metadata-calamine-head-to-head.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-metadata-calamine-clean-b6925afe/metadata-calamine-head-to-head.json --json --metric medianMs --require-profile xlsx-read-sota --assert-profile-leader ascend > /private/tmp/ascend-metadata-calamine-clean-b6925afe/metadata-calamine-scoreboard.json
```

Environment:

- Commit: `b6925afe`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-metadata-calamine-clean-b6925afe`
- Bun: `1.3.13`
- Python: `3.9.6`
- Runtime profile: `category read`, `executionScope external-process`,
  `workload metadata-only`, `readSource raw-ooxml`, `validationMode each`,
  `repeat 15`, `warmup 3`.

Result:

| Runner | Status | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `python-calamine-metadata-only` | ran/won | 0.056 | 0.078 | 0.122 | 28.4 MiB | `metadataOnlyRead: true`, `cellsHydrated: false`, three sheets loaded |
| `sheetjs-metadata-only` | ran/lost vs Calamine, ran/won vs Ascend median | 2.864 | 9.810 | 0.714 | 126.7 MiB | `metadataOnlyRead: true`, `cellsHydrated: false`, three sheets loaded |
| `ascend-external-metadata-only` | ran/lost | 3.211 | 8.596 | 0.623 | 88.4 MiB | `metadataOnlyRead: true`, `cellsHydrated: false`, three sheets loaded |
| `openpyxl-metadata-only` | ran/lost | 9.631 | 16.499 | 0.286 | 52.4 MiB | `metadataOnlyRead: true`, `cellsHydrated: false`, three sheets loaded |

Gate result: `profileLeaderFailures` contains the metadata-only loss to
`python-calamine-metadata-only`. ExcelJS, Apache POI, and ClosedXML remain
`unsupported-operation` coverage gaps for this profile row and are not wins.

Rejected optimization: a current-worktree patch that skipped plain worksheet
relationship probes in metadata-only mode preserved focused semantic tests, but
paired local `readXlsx(..., { mode: "metadata-only" })` measurement did not
validate it: baseline median `0.105 ms`, patched median `0.168 ms` on the same
generated input after warmup. The patch was reverted.

Humble allowed wording:

> Ascend has comparable metadata-only external evidence, but python-calamine is
> faster on the generated metadata-only lane. This row is an optimization target,
> not an Ascend speed win.

Forbidden wording:

- "Ascend is fastest for metadata-only XLSX reads."
- "Ascend beats Calamine on metadata-only open."
- Any wording that counts unsupported metadata-only competitors as wins.

Next action: optimize only if profiling identifies a production change that
reduces Ascend's measured metadata-only open time without weakening active
content inventory or document-property inspection. Otherwise keep the
metadata-only claim downgraded.

## Cycle: Current Metadata-Only Calamine Recheck

Classification: kill/defer. The current clean rerun keeps
`python-calamine-metadata-only` as the metadata-only median winner, and the
only narrow production candidate tested in this block did not improve Ascend.
No production optimization is carried forward.

Workflow: generated XLSX metadata-only open/inspect, loading workbook and sheet
metadata without hydrating cells.

Why it matters for release: this is the measured read workflow where Ascend is
not at `ran/won`. It gates any metadata-only speed wording used for safe unknown
workbook inspection before edit planning.

Public/tracked-clean input: `competitive-io` generated `metadata-only`
`raw-ooxml`, 200 rows x 20 columns, three workbook sheets, 15,347 input bytes,
from tracked benchmark code at commit `c70385bc`.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-metadata-current-c70385bc c70385bc
cd /private/tmp/ascend-metadata-current-c70385bc
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-metadata-current-c70385bc-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only,sheetjs-metadata-only,openpyxl-metadata-only,python-calamine-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 15 --warmup 3 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-current-c70385bc-runs/metadata-calamine-head-to-head.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-metadata-current-c70385bc-runs/metadata-calamine-head-to-head.json --json --metric medianMs --require-profile xlsx-read-sota --assert-profile-leader ascend > /private/tmp/ascend-metadata-current-c70385bc-runs/metadata-calamine-scoreboard.json
```

Rejected optimization command:

```bash
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only,python-calamine-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 15 --warmup 3 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-current-c70385bc-patched-runs/metadata-calamine-head-to-head.json
```

Environment:

- Commit: `c70385bc30a35aaf208b98e1566de724666d7d43`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-metadata-current-c70385bc`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Python runners: `python3` from the command `PATH`
- Runtime profile: `category read`, `executionScope external-process`,
  `workload metadata-only`, `readSource raw-ooxml`, `validationMode each`,
  `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-metadata-current-c70385bc-runs/metadata-calamine-head-to-head.json
/private/tmp/ascend-metadata-current-c70385bc-runs/metadata-calamine-scoreboard.json
/private/tmp/ascend-metadata-current-c70385bc-patched-runs/metadata-calamine-head-to-head.json
```

Result, repeat 15 after 3 warmups:

| Runner | Status | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `python-calamine-metadata-only` | ran/won | 0.070 | 0.151 | 0.309 | 28.4 MiB | Comparable for the generated plain sheet-list metadata contract: three sheets loaded, cells not hydrated. |
| `ascend-external-metadata-only` | ran/lost vs Calamine, ran/won vs SheetJS/openpyxl | 1.208 | 6.171 | 0.840 | 83.4 MiB | Same plain metadata contract plus Ascend's SDK metadata-only inspection surface. |
| `sheetjs-metadata-only` | ran/lost | 1.643 | 5.580 | 0.675 | 151.5 MiB | Same plain metadata contract. |
| `openpyxl-metadata-only` | ran/lost | 5.190 | 7.246 | 0.271 | 53.4 MiB | Same plain metadata contract. |

Scoreboard result: `profileLeaderFailures` contains
`winner=python-calamine-metadata-only expected=ascend` for
`read-metadata-only`.

Rejected optimization: a current-worktree patch skipped preservation-capsule
construction when metadata-only inputs had no unconsumed package parts beyond
already parsed document properties. Focused correctness tests passed for
metadata-only sheet structure, document properties, and active-content inventory,
but the paired benchmark did not validate a win: patched Ascend median
`1.348 ms`, p95 `4.410 ms`, CV `0.628`, peak RSS `90.6 MiB`, while the clean
current baseline was `1.208 ms`, p95 `6.171 ms`, CV `0.840`, peak RSS
`83.4 MiB`. The patch was reverted.

Semantic boundary: this row is comparable only for the generated plain
metadata-only sheet-list contract. It is not evidence that Calamine supports
Ascend's safe-open metadata-only release workflow for document properties,
active content inventory, package risk reporting, or edit-planning trust
decisions.

Humble allowed wording:

> On the generated plain metadata-only workload, Calamine is faster than Ascend.
> Ascend remains faster than SheetJS and OpenPyXL in this rerun, but the
> Calamine head-to-head is a real measured loss for the plain sheet-list
> metadata contract and is not a promoted speed claim.

Forbidden wording:

- "Ascend is fastest for metadata-only XLSX reads."
- "Ascend beats Calamine on metadata-only open."
- "Calamine proves faster safe-open trust inspection than Ascend."
- Any wording that treats unsupported safe-open metadata semantics as a speed win.

Next action: kill the capsule-skip optimization target and defer further
metadata-only production work until profiling identifies a narrower measured
cost center. Keep the metadata-only claim downgraded.

## Cycle: Current Metadata-Only Calamine Recheck at `cc689bcc`

Classification: kill/defer. A clean current rerun keeps
`python-calamine-metadata-only` as the winner for the comparable plain
metadata-only sheet-list contract. A narrow file-backed path-open candidate was
tested and killed because it did not beat the clean current Ascend baseline.

Workflow: generated XLSX metadata-only open/inspect, loading workbook and sheet
metadata without hydrating cells.

Why it matters for release: metadata-only open is the first unknown-workbook
inspection mode before edit planning. Any speed wording for this workflow must
survive the native Calamine sheet-list baseline and must not count unsupported
safe-open semantics as wins.

Public/tracked-clean input: `competitive-io` generated `metadata-only`
`raw-ooxml`, 200 rows x 20 columns, three workbook sheets, 15,347 input bytes,
from tracked benchmark code at commit `cc689bcc`.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-metadata-current-cc689bcc cc689bccef33
cd /private/tmp/ascend-metadata-current-cc689bcc
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-metadata-current-cc689bcc-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only,sheetjs-metadata-only,openpyxl-metadata-only,python-calamine-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 15 --warmup 3 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-current-cc689bcc-runs/metadata-calamine-head-to-head.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-metadata-current-cc689bcc-runs/metadata-calamine-head-to-head.json --json --metric medianMs --require-profile xlsx-read-sota --assert-profile-leader ascend > /private/tmp/ascend-metadata-current-cc689bcc-runs/metadata-calamine-scoreboard.json
```

Rejected optimization command:

```bash
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only,python-calamine-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 15 --warmup 3 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-current-cc689bcc-patched-runs/metadata-calamine-head-to-head.json
```

Environment:

- Commit: `cc689bccef337d2cf2ef9a43158258e3e7b09fca`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-metadata-current-cc689bcc`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Python runners: `python3` from the command `PATH`
- Runtime profile: `category read`, `executionScope external-process`,
  `workload metadata-only`, `readSource raw-ooxml`, `validationMode each`,
  `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-metadata-current-cc689bcc-runs/metadata-calamine-head-to-head.json
/private/tmp/ascend-metadata-current-cc689bcc-runs/metadata-calamine-scoreboard.json
/private/tmp/ascend-metadata-current-cc689bcc-patched-runs/metadata-calamine-head-to-head.json
```

Result, repeat 15 after 3 warmups:

| Runner | Status | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `python-calamine-metadata-only` | ran/won | 0.045 | 0.095 | 0.285 | 28.5 MiB | Same timing lane and assertions: `metadataOnlyRead: true`, `cellsHydrated: false`, three sheets loaded. |
| `ascend-external-metadata-only` | ran/lost vs Calamine, ran/won vs SheetJS/openpyxl | 0.450 | 0.884 | 0.301 | 91.4 MiB | Same plain sheet-list/no-cell-hydration contract through Ascend's SDK metadata-only open surface. |
| `sheetjs-metadata-only` | ran/lost | 0.692 | 0.992 | 0.159 | 150.6 MiB | Same plain sheet-list/no-cell-hydration contract using SheetJS `bookSheets`. |
| `openpyxl-metadata-only` | ran/lost | 2.317 | 4.292 | 0.274 | 53.5 MiB | Same plain sheet-list/no-cell-hydration contract using read-only OpenPyXL metadata inventory. |

Scoreboard result: `profileLeaderFailures` contains
`winner=python-calamine-metadata-only expected=ascend` for
`read-metadata-only`. ExcelJS, Apache POI, and ClosedXML remain
`unsupported-operation` coverage gaps for metadata-only and are not wins.

Rejected optimization: a current-worktree patch extended the SDK's existing
file-backed ZIP path to `Ascend.open(path, { mode: "metadata-only" })`, with a
focused test proving metadata-only path opens still inventory `vbaProject` and
`vbaSignature` sidecars. The paired benchmark did not validate a win: patched
Ascend median `0.550 ms`, p95 `0.928 ms`, CV `0.225`, peak RSS `91.1 MiB`,
while the clean current baseline was `0.450 ms`, p95 `0.884 ms`, CV `0.301`,
peak RSS `91.4 MiB`. The patch was reverted.

Semantic boundary: this is comparable for the generated plain metadata-only
sheet-list contract only. Calamine does not prove parity with Ascend's broader
safe-open trust inspection for document properties, active-content inventory,
package risk reporting, or edit-planning decisions in the existing runner.

Humble allowed wording:

> On the generated plain metadata-only workload at `cc689bcc`, Calamine is
> faster than Ascend for the comparable sheet-list/no-cell-hydration timing
> lane. Ascend remains faster than SheetJS and OpenPyXL in this rerun, but this
> row is a measured Calamine loss, not a speed claim.

Forbidden wording:

- "Ascend is fastest for metadata-only XLSX reads."
- "Ascend beats Calamine on metadata-only open."
- "Calamine proves faster safe-open trust inspection than Ascend."
- Any wording that treats unsupported metadata-only runners or unsupported
  safe-open semantics as wins.

Next action: kill the file-backed metadata-only path-open target and keep the
metadata-only speed claim downgraded. Do not spend more production work on this
row until profiling identifies a narrower cost center that can plausibly close
the Calamine gap without weakening active-content or document-property
inspection.

## Cycle: Dense Values Write SOTA Gate

Classification: comparable external evidence plus defer. The first full
external dense-values write row reported a `profileLeaderFailures` loss to
`rust-xlsxwriter`, but Ascend's row was very noisy. A repeat-15 rerun against
the fastest comparable writers did not reproduce the loss: Ascend was the median
winner. No production optimization is justified from this evidence.

Workflow: generated XLSX write for dense numeric values, 2000 rows x 20 columns.

Why it matters for release: generated value writes are the basic commit/export
path after an agent produces or rewrites a workbook. This is the first `xlsx-write-sota`
head-to-head row that can block write-speed wording.

Public/tracked-clean input: `competitive-io` generated the `dense-values`
`source-mode generated-write` workload from tracked benchmark code at commit
`4b8b82b6`. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-dense-current-4b8b82b6 4b8b82b6
cd /private/tmp/ascend-write-dense-current-4b8b82b6
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-dense-current-4b8b82b6-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,xlsxwriter,xlsxwriter-constant-memory,pyexcelerate,pyexcelerate-range,pyexcelerate-cell,openpyxl,openpyxl-write-only,apache-poi,closedxml,rust-xlsxwriter,excelize,fastexcel-java --workload dense-values --repeat 5 --warmup 1 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-head-to-head.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-head-to-head.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,fastexcel-java --workload dense-values --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `4b8b82b602abc2980138d670f5e6591199beb39a`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-dense-current-4b8b82b6`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload dense-values`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-head-to-head.json
/private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-scoreboard.json
/private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-fastest-repeat15.json
/private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-fastest-repeat15-scoreboard.json
```

Full external row, repeat 5 after 1 warmup:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/lost | 70.080 | 222.688 | 0.674 | 58.8 MiB | 172260 |
| `rust-xlsxwriter` | ran/won vs Ascend | 26.806 | 31.020 | 0.076 | 17.3 MiB | 119133 |
| `excelize` | ran/won vs Ascend | 64.487 | 80.300 | 0.226 | 17.0 MiB | 120015 |
| `fastexcel-java` | ran/won vs Ascend | 65.433 | 187.666 | 0.657 | 360.0 MiB | 161599 |
| `sheetjs` | ran/lost vs Ascend | 89.976 | 102.173 | 0.070 | 212.8 MiB | 1181431 |
| `pyexcelerate-cell` | ran/lost vs Ascend | 276.001 | 986.370 | 0.863 | 60.7 MiB | 116623 |
| `pyexcelerate` | ran/lost vs Ascend | 351.764 | 436.320 | 0.149 | 60.8 MiB | 116623 |
| `apache-poi` | ran/lost vs Ascend | 360.049 | 840.329 | 0.499 | 888.0 MiB | 124020 |
| `exceljs` | ran/lost vs Ascend | 376.254 | 531.876 | 0.193 | 204.5 MiB | 121315 |
| `pyexcelerate-range` | ran/lost vs Ascend | 392.131 | 833.361 | 0.441 | 63.0 MiB | 116623 |
| `openpyxl` | ran/lost vs Ascend | 401.464 | 521.731 | 0.164 | 99.0 MiB | 124665 |
| `xlsxwriter` | ran/lost vs Ascend | 422.958 | 468.257 | 0.114 | 71.4 MiB | 118761 |
| `openpyxl-write-only` | ran/lost vs Ascend | 448.541 | 468.733 | 0.064 | 82.4 MiB | 124636 |
| `xlsxwriter-constant-memory` | ran/lost vs Ascend | 468.404 | 728.297 | 0.275 | 53.4 MiB | 118249 |
| `closedxml` | runner unavailable | n/a | n/a | n/a | n/a | n/a |

Focused fastest-writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 4.851 | 9.392 | 0.244 | 73.7 MiB | 172260 |
| `excelize` | ran/lost vs Ascend | 17.403 | 18.270 | 0.033 | 17.6 MiB | 120015 |
| `fastexcel-java` | ran/lost vs Ascend | 21.152 | 42.718 | 0.314 | 576.0 MiB | 161598 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 31.758 | 36.138 | 0.070 | 18.8 MiB | 119134 |

Scoreboard result:

- Full external row: `profileLeaderFailures` contains
  `winner=rust-xlsxwriter expected=ascend`.
- Focused repeat-15 fastest-writer rerun: `profileLeaderFailures: []`.
- `closedxml` was `runner unavailable` because the .NET build failed with
  `CSSM_ModuleLoad()`. It is not counted as a win.

Semantic comparability: all passing rows write the same generated dense numeric
sheet and pass external post-write semantic validation. File-size and memory are
not equal: Ascend emits a larger XLSX than the fastest native writers and uses
more RSS than Excelize/rust_xlsxwriter in the repeat-15 rerun.

Humble allowed wording:

> On the generated 2000 x 20 dense-values write row, a noisy full external
> repeat-5 run showed native writers ahead of Ascend, but a focused repeat-15
> rerun against the fastest comparable writers had Ascend as the median winner.
> This is scoped write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every generated XLSX writer."
- "Ascend beats ClosedXML on dense-value writes."
- "Ascend produces the smallest dense-value XLSX."

Next action: defer production optimization from this row. Continue write-profile
coverage with the next release workflow only after a clean multi-workload
`xlsx-write-sota` gate identifies a durable leader failure.

## Cycle: Dense Values Write Current Fastest Comparable Row

Classification: comparable external evidence. This row refreshes the
dense-values write baseline at current `HEAD` after the small dense streaming
fallback optimization. It is a scoped current-commit win, not a broad
`xlsx-write-sota` claim.

Workflow: generated XLSX write for dense numeric values, 2000 rows x 20 columns.

Why it matters for release: dense value export is the simplest generated
commit/write path after an agent creates or rewrites workbook values. It is also
the write row most directly affected by the dense streaming fallback.

Public/tracked-clean input: `competitive-io` generated the `dense-values`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `905ecb5e`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-dense-current-905ecb5e 905ecb5e
cd /private/tmp/ascend-write-dense-current-905ecb5e
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-dense-current-905ecb5e-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,fastexcel-java --workload dense-values --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-dense-current-905ecb5e-runs/write-dense-values-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-dense-current-905ecb5e-runs/write-dense-values-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-dense-current-905ecb5e-runs/write-dense-values-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `905ecb5e1754fd491d5e9f4b41ddcaf1d6a428e1`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-dense-current-905ecb5e`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64, macOS `26.4`, kernel
  `25.4.0` on `RELEASE_ARM64_T6041`
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload dense-values`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-dense-current-905ecb5e-runs/write-dense-values-fastest-repeat15.json
/private/tmp/ascend-write-dense-current-905ecb5e-runs/write-dense-values-fastest-repeat15-scoreboard.json
```

Focused fastest comparable writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 6.082 | 10.191 | 0.216 | 78.2 MiB | 172259 |
| `fastexcel-java` | ran/lost vs Ascend | 27.540 | 48.041 | 0.268 | 504.0 MiB | 161599 |
| `excelize` | ran/lost vs Ascend | 34.534 | 49.529 | 0.194 | 17.8 MiB | 120015 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 71.640 | 134.871 | 0.290 | 23.0 MiB | 119133 |

Scoreboard result:

- Focused repeat-15 fastest-writer row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The scoreboard command exits nonzero for the full profile because this row is
  not full `xlsx-write-sota` coverage. Missing/omitted full-profile libraries,
  unsupported rows, and blocked runners are not counted as wins.

Semantic comparability: all listed rows reopened successfully, matched the
expected one-sheet and 40,000-cell shape, and passed semantic cell value hash
validation. Ascend and Excelize also matched ordered semantic cell hashes.
FastExcel Java and rust_xlsxwriter passed sorted semantic value equality but did
not match ordered semantic value hashes, so their speed rows are useful
lower-fidelity value-write comparisons, not byte/order-equivalent output claims.
Ascend is faster in this row but uses more RSS and emits a larger XLSX than
Excelize and rust_xlsxwriter.

Humble allowed wording:

> On the generated 2000 x 20 dense-values write row at commit `905ecb5e`,
> Ascend's focused external repeat-15 run had the fastest median and p95 among
> the completed fastest comparable writers in this row. This is scoped generated
> dense-value write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every generated XLSX writer."
- "Ascend beats omitted, unsupported, or blocked dense-value writers."
- "Ascend produces the smallest dense-value XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: defer production optimization for dense-values. Continue with the
next priority workflow only if it can produce a comparable baseline row,
validated optimization, or explicit claim downgrade.

## Cycle: Plain Text Write SOTA Gate

Classification: comparable external evidence plus defer. The clean
external-process `plain-text` write row has no leader failure among passing
comparable writers: Ascend is the median winner. ClosedXML failed during its
.NET build and is recorded as runner unavailable, not as a win. No production
optimization is justified from this evidence.

Workflow: generated XLSX write for plain text values, 2000 rows x 20 columns.

Why it matters for release: text-heavy generated exports are a common
agent-produced workbook output, and `plain-text` is one of the value-write rows
required by the existing `xlsx-write-sota` profile.

Public/tracked-clean input: `competitive-io` generated the `plain-text`
`source-mode generated-write` workload from tracked benchmark code at commit
`98752c84`. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-plain-text-current-98752c84 98752c84
cd /private/tmp/ascend-write-plain-text-current-98752c84
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-plain-text-current-98752c84-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,xlsxwriter,xlsxwriter-constant-memory,pyexcelerate,pyexcelerate-range,pyexcelerate-cell,openpyxl,openpyxl-write-only,apache-poi,closedxml,rust-xlsxwriter,excelize,fastexcel-java --workload plain-text --repeat 5 --warmup 1 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-plain-text-current-98752c84-runs/write-plain-text-head-to-head.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-plain-text-current-98752c84-runs/write-plain-text-head-to-head.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-plain-text-current-98752c84-runs/write-plain-text-scoreboard.json
```

Environment:

- Commit: `98752c84e93e0058d6d646138fb1b1c17ebfa389`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-plain-text-current-98752c84`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload plain-text`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-plain-text-current-98752c84-runs/write-plain-text-head-to-head.json
/private/tmp/ascend-write-plain-text-current-98752c84-runs/write-plain-text-scoreboard.json
```

External row, repeat 5 after 1 warmup:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 16.937 | 29.104 | 0.313 | 64.4 MiB | 169099 |
| `excelize` | ran/lost vs Ascend | 28.623 | 39.311 | 0.197 | 21.2 MiB | 142890 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 39.684 | 40.941 | 0.066 | 24.1 MiB | 229138 |
| `sheetjs` | ran/lost vs Ascend | 49.825 | 62.671 | 0.119 | 244.0 MiB | 1832541 |
| `fastexcel-java` | ran/lost vs Ascend | 57.577 | 80.924 | 0.173 | 680.0 MiB | 227254 |
| `pyexcelerate` | ran/lost vs Ascend | 125.436 | 141.870 | 0.091 | 72.7 MiB | 144372 |
| `pyexcelerate-range` | ran/lost vs Ascend | 128.641 | 141.047 | 0.062 | 74.3 MiB | 144372 |
| `xlsxwriter-constant-memory` | ran/lost vs Ascend | 131.639 | 144.159 | 0.060 | 59.4 MiB | 141061 |
| `xlsxwriter` | ran/lost vs Ascend | 135.849 | 149.731 | 0.048 | 87.5 MiB | 230174 |
| `exceljs` | ran/lost vs Ascend | 137.375 | 161.997 | 0.106 | 263.7 MiB | 232106 |
| `pyexcelerate-cell` | ran/lost vs Ascend | 147.762 | 886.007 | 1.127 | 74.4 MiB | 144372 |
| `openpyxl-write-only` | ran/lost vs Ascend | 177.626 | 200.032 | 0.073 | 92.3 MiB | 140903 |
| `openpyxl` | ran/lost vs Ascend | 246.598 | 290.885 | 0.134 | 96.2 MiB | 140929 |
| `apache-poi` | ran/lost vs Ascend | 422.230 | 728.093 | 0.302 | 952.0 MiB | 229438 |
| `closedxml` | runner unavailable | n/a | n/a | n/a | n/a | n/a |

Scoreboard result:

- `profileLeaderFailures: []`
- The full `xlsx-write-sota` gate still fails coverage because this is a single
  row and because `closedxml` is ineligible for `plain-text`.
- `closedxml` failed with `CSSM_ModuleLoad()` during the .NET build. It is not
  counted as a win.

Semantic comparability: all passing rows write the same generated plain-text
sheet and pass external post-write semantic validation for one sheet and 40,000
cells. File-size and memory are not equal: Ascend is not the smallest output and
uses more RSS than Excelize, rust_xlsxwriter, and constant-memory XlsxWriter.

Humble allowed wording:

> On the generated 2000 x 20 plain-text write row, Ascend was the median winner
> among passing comparable external writers. This is scoped value-write evidence,
> not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every generated XLSX writer."
- "Ascend beats ClosedXML on plain-text writes."
- "Ascend produces the smallest plain-text XLSX."

Next action: defer production optimization from this row and continue only with
another existing `xlsx-write-sota` row or a measured workflow loss.

## Cycle: String Heavy Write SOTA Gate

Classification: comparable external evidence plus defer. The first full
external `string-heavy` write row had Ascend behind rust_xlsxwriter, Excelize,
SheetJS, and FastExcel Java by median, but the row was noisy across several
runners. A repeat-15 rerun against the fastest comparable writers reversed the
median result: Ascend was the median winner. The tail remains noisy, so this is
not a production optimization target or broad write-speed claim.

Workflow: generated XLSX write for varied string values, 2000 rows x 20 columns.

Why it matters for release: string-heavy generated exports exercise a common
agent output shape for reports, lists, labels, regions, customer notes, and
statuses. This row is one of the value-write workloads required by the existing
`xlsx-write-sota` profile.

Public/tracked-clean input: `competitive-io` generated the `string-heavy`
`source-mode generated-write` workload from tracked benchmark code at commit
`67b900ed`. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-string-heavy-current-67b900ed 67b900ed
cd /private/tmp/ascend-write-string-heavy-current-67b900ed
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-string-heavy-current-67b900ed-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,xlsxwriter,xlsxwriter-constant-memory,pyexcelerate,pyexcelerate-range,pyexcelerate-cell,openpyxl,openpyxl-write-only,apache-poi,closedxml,rust-xlsxwriter,excelize,fastexcel-java --workload string-heavy --repeat 5 --warmup 1 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-head-to-head.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-head-to-head.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-scoreboard.json
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,sheetjs,fastexcel-java --workload string-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `67b900edffc46955d9bbc8f98782600facb83ec2`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-string-heavy-current-67b900ed`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload string-heavy`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-head-to-head.json
/private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-scoreboard.json
/private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-fastest-repeat15.json
/private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-fastest-repeat15-scoreboard.json
```

Full external row, repeat 5 after 1 warmup:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/lost | 152.628 | 209.349 | 0.566 | 60.7 MiB | 201985 |
| `rust-xlsxwriter` | ran/won vs Ascend | 65.802 | 135.550 | 0.424 | 25.5 MiB | 237837 |
| `excelize` | ran/won vs Ascend | 75.903 | 186.230 | 0.540 | 21.8 MiB | 218447 |
| `sheetjs` | ran/won vs Ascend | 117.119 | 328.740 | 0.611 | 173.3 MiB | 2016032 |
| `fastexcel-java` | ran/won vs Ascend | 138.447 | 188.627 | 0.180 | 688.0 MiB | 260427 |
| `exceljs` | ran/lost vs Ascend | 282.215 | 355.848 | 0.151 | 206.7 MiB | 240319 |
| `pyexcelerate` | ran/lost vs Ascend | 363.745 | 624.459 | 0.341 | 67.2 MiB | 217988 |
| `pyexcelerate-cell` | ran/lost vs Ascend | 443.577 | 610.712 | 0.327 | 72.2 MiB | 217988 |
| `xlsxwriter` | ran/lost vs Ascend | 467.612 | 792.297 | 0.444 | 79.9 MiB | 238362 |
| `xlsxwriter-constant-memory` | ran/lost vs Ascend | 471.814 | 646.756 | 0.197 | 58.3 MiB | 211656 |
| `pyexcelerate-range` | ran/lost vs Ascend | 517.739 | 2124.323 | 0.803 | 64.4 MiB | 217988 |
| `apache-poi` | ran/lost vs Ascend | 675.989 | 2551.732 | 0.809 | 1176.0 MiB | 241513 |
| `openpyxl-write-only` | ran/lost vs Ascend | 784.534 | 923.160 | 0.169 | 84.6 MiB | 211499 |
| `openpyxl` | ran/lost vs Ascend | 1408.501 | 1990.245 | 0.353 | 88.0 MiB | 211519 |
| `closedxml` | runner unavailable | n/a | n/a | n/a | n/a | n/a |

Focused fastest-writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 63.896 | 372.935 | 0.927 | 65.5 MiB | 201985 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 89.210 | 175.070 | 0.322 | 23.9 MiB | 237837 |
| `sheetjs` | ran/lost vs Ascend | 94.161 | 209.940 | 0.352 | 276.5 MiB | 2016032 |
| `fastexcel-java` | ran/lost vs Ascend | 112.387 | 512.219 | 0.770 | 1544.0 MiB | 260427 |
| `excelize` | ran/lost vs Ascend | 229.576 | 334.389 | 0.351 | 24.3 MiB | 218447 |

Scoreboard result:

- Full external row: group winner was `rust-xlsxwriter`; the single-row
  scoreboard still reports `profileLeaderFailures: []` because full
  `xlsx-write-sota` coverage is missing.
- Focused repeat-15 fastest-writer rerun: group winner was
  `ascend-external-writer`; `profileLeaderFailures: []`.
- `closedxml` failed with `CSSM_ModuleLoad()` during the .NET build. It is not
  counted as a win.

Semantic comparability: all passing rows write the same generated string-heavy
sheet and pass external post-write semantic validation for one sheet and 40,000
cells. Memory and size tradeoffs remain material: rust_xlsxwriter and Excelize
use less RSS; SheetJS emits a much larger file; Ascend's focused rerun median
wins but its p95 and CV are noisy.

Humble allowed wording:

> On the generated 2000 x 20 string-heavy write row, a noisy full repeat-5 run
> showed native writers ahead of Ascend, but a focused repeat-15 rerun against
> the fastest comparable writers had Ascend as the median winner. This is scoped
> value-write evidence with a noisy tail, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend always beats rust_xlsxwriter, Excelize, SheetJS, or FastExcel Java on
  string-heavy writes."
- "Ascend beats ClosedXML on string-heavy writes."
- "Ascend has the best tail latency for string-heavy writes."
- "Ascend produces the smallest string-heavy XLSX."

Next action: defer production optimization from this row. If string-heavy
matters for a release claim later, attack the noisy Ascend tail with profiling
before changing writer code.

## Cycle: String Heavy Write Optimization

Classification: validated optimization. The noisy `string-heavy` write gap was
real under the benchmark runtime, not under the default shell Bun. The external
matrix PATH uses `/Users/arjun/.bun/bin/bun` `1.3.13`; under that runtime the
small dense streaming ZIP path was much slower and noisier than the buffered
dense writer. Commit `bd162937` routes small dense streaming writes through the
buffered dense ZIP path, preserving the streaming API while avoiding async
stream overhead for small generated sheets.

Workflow: generated XLSX write for varied string values, 2000 rows x 20 columns.

Why it matters for release: this directly improves the generated string export
workflow that previously had a repeat-5 median loss and a noisy repeat-15 tail.
It also improves small dense generated numeric/text writes because they share
the same dense streaming writer path.

Public/tracked-clean input: `competitive-io` generated the `string-heavy`
`source-mode generated-write` workload from tracked benchmark code at commit
`bd162937`. No private corpus or local research workbook was used.

Production change:

- `packages/io-xlsx/src/writer/dense-rows.ts` now uses the synchronous dense
  ZIP builder when the estimated dense sheet XML is at or below 4 MiB.
- `packages/io-xlsx/src/writer/writer.test.ts` proves small dense streaming
  output is byte-equivalent to the buffered dense writer and reopens with the
  expected values.

Commands:

```bash
bun test packages/io-xlsx/src/writer/writer.test.ts -t "small dense streaming output"
bun test packages/io-xlsx/src/writer/writer.test.ts -t "dense rows"
bun test fixtures/benchmarks/competitive-io.test.ts -t "string-heavy"
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/xlsx-write-phase.ts --workload string-heavy --rows 2000 --cols 20 --repeat 12 --warmup 3 --streaming --json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/xlsx-write-phase.ts --workload dense-values --rows 2000 --cols 20 --repeat 12 --warmup 3 --streaming --json
git worktree add --detach /private/tmp/ascend-write-string-heavy-optimized-bd162937 bd162937
cd /private/tmp/ascend-write-string-heavy-optimized-bd162937
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-string-heavy-optimized-bd162937-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,sheetjs,fastexcel-java --workload string-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-string-heavy-optimized-bd162937-runs/write-string-heavy-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-string-heavy-optimized-bd162937-runs/write-string-heavy-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-string-heavy-optimized-bd162937-runs/write-string-heavy-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `bd1629373a9a4a17edd2db9125b6cd19e6df7504`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-string-heavy-optimized-bd162937`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload string-heavy`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-string-heavy-optimized-bd162937-runs/write-string-heavy-fastest-repeat15.json
/private/tmp/ascend-write-string-heavy-optimized-bd162937-runs/write-string-heavy-fastest-repeat15-scoreboard.json
```

Phase evidence under Bun `1.3.13`:

| Workload | Before median write ms | After median write ms | Before tail/noise | After tail/noise |
| --- | ---: | ---: | --- | --- |
| `string-heavy` dense streaming | 107.176 | 12.418 | samples up to 198.748 ms | samples up to 31.683 ms |
| `dense-values` dense streaming | 76.335 | 11.179 | samples up to 176.338 ms | samples up to 24.411 ms |

Optimized focused fastest-writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 5.001 | 6.077 | 0.099 | 91.8 MiB | 201984 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 30.237 | 34.996 | 0.079 | 24.3 MiB | 237837 |
| `fastexcel-java` | ran/lost vs Ascend | 30.614 | 44.745 | 0.183 | 808.0 MiB | 260426 |
| `excelize` | ran/lost vs Ascend | 33.528 | 41.901 | 0.108 | 23.3 MiB | 218447 |
| `sheetjs` | ran/lost vs Ascend | 62.450 | 224.709 | 0.622 | 278.0 MiB | 2016032 |

Scoreboard result:

- Focused repeat-15 fastest-writer rerun: group winner was
  `ascend-external-writer`; `profileLeaderFailures: []`.
- The full `xlsx-write-sota` gate still fails coverage because this is a
  focused row, not a full-profile promotion run.

Semantic comparability: all passing rows write the same generated string-heavy
sheet and pass external post-write semantic validation for one sheet and 40,000
cells. Memory and size tradeoffs remain material: rust_xlsxwriter and Excelize
use less RSS, while Ascend now wins median and p95 on this focused row.

Humble allowed wording:

> On the generated 2000 x 20 string-heavy write row, after `bd162937`, Ascend's
> focused external repeat-15 row was faster by median and p95 than
> rust_xlsxwriter, FastExcel Java, Excelize, and SheetJS. This is scoped
> generated string-write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every generated XLSX writer."
- "Ascend beats ClosedXML or NPOI on string-heavy writes."
- "Ascend has the smallest string-heavy XLSX."

Next action: continue optimizing or bounding the next existing `xlsx-write-sota`
gap; do not revisit string-heavy unless a full-profile rerun regresses this row.

## Cycle: String Heavy Write Current Fastest Comparable Row

Classification: comparable external evidence. This row refreshes the optimized
string-heavy write result at current `HEAD`. It confirms the previous production
optimization remains a scoped current-commit win with low noise, not a broad
`xlsx-write-sota` claim.

Workflow: generated XLSX write for varied string values, 2000 rows x 20 columns.

Why it matters for release: string-heavy generated exports are common in
agent-produced reports and tables. This row was previously the measured write
gap that justified the dense streaming fallback optimization, so current
release evidence should prove it stayed fixed.

Public/tracked-clean input: `competitive-io` generated the `string-heavy`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `9df35fd6`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-string-current-9df35fd6 9df35fd6
cd /private/tmp/ascend-write-string-current-9df35fd6
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-string-current-9df35fd6-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,sheetjs,fastexcel-java --workload string-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-string-current-9df35fd6-runs/write-string-heavy-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-string-current-9df35fd6-runs/write-string-heavy-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-string-current-9df35fd6-runs/write-string-heavy-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `9df35fd62cb07ed0c790894bf2a0e008f7a19174`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-string-current-9df35fd6`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64, macOS `26.4`, kernel
  `25.4.0` on `RELEASE_ARM64_T6041`
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload string-heavy`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-string-current-9df35fd6-runs/write-string-heavy-fastest-repeat15.json
/private/tmp/ascend-write-string-current-9df35fd6-runs/write-string-heavy-fastest-repeat15-scoreboard.json
```

Focused fastest comparable writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 6.538 | 7.515 | 0.061 | 92.1 MiB | 201984 |
| `fastexcel-java` | ran/lost vs Ascend | 34.463 | 48.012 | 0.189 | 904.0 MiB | 260427 |
| `excelize` | ran/lost vs Ascend | 35.803 | 40.078 | 0.065 | 22.0 MiB | 218447 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 49.872 | 100.163 | 0.264 | 24.4 MiB | 237838 |
| `sheetjs` | ran/lost vs Ascend | 53.164 | 60.661 | 0.059 | 279.2 MiB | 2016032 |

Scoreboard result:

- Focused repeat-15 fastest-writer row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The scoreboard command exits nonzero for the full profile because this row is
  not full `xlsx-write-sota` coverage. Missing/omitted full-profile libraries,
  unsupported rows, and blocked runners are not counted as wins.

Semantic comparability: all listed rows reopened successfully, matched the
expected one-sheet and 40,000-cell shape, and passed semantic cell value hash
validation. Ascend, Excelize, and SheetJS also matched ordered semantic cell
hashes. FastExcel Java and rust_xlsxwriter passed sorted semantic value equality
but did not match ordered semantic value hashes, so their speed rows are useful
lower-fidelity value-write comparisons, not byte/order-equivalent output claims.
Ascend is faster in this row but uses more RSS than Excelize and
rust_xlsxwriter, and does not emit the smallest XLSX.

Humble allowed wording:

> On the generated 2000 x 20 string-heavy write row at commit `9df35fd6`,
> Ascend's focused external repeat-15 run had the fastest median and p95 among
> the completed fastest comparable writers in this row. This is scoped generated
> string-write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every generated XLSX writer."
- "Ascend beats omitted, unsupported, or blocked string-heavy writers."
- "Ascend produces the smallest string-heavy XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: defer production optimization for string-heavy. Continue with the
next priority workflow only if it can produce a comparable baseline row,
validated optimization, or explicit claim downgrade.

## Cycle: String Heavy Write Current Fastest Comparable Row at `bea0d001`

Classification: comparable external evidence plus defer. This refreshes the
optimized string-heavy write row at current `HEAD` after the latest release
claim commits. Ascend is the median and p95 winner among the completed fastest
comparable writers in this row. The Ascend CV is noisy, but the tail remains
below the compared external writers, so no production optimization is justified
from this run.

Workflow: generated XLSX write for varied string values, 2000 rows x 20 columns.

Why it matters for release: string-heavy generated exports are common in
agent-produced reports, lists, labels, notes, and statuses. This row previously
exposed the dense streaming write gap, so current release evidence should prove
the optimized path remains externally defensible.

Public/tracked-clean input: `competitive-io` generated the `string-heavy`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `bea0d001`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-string-current-bea0d001 bea0d001e778
cd /private/tmp/ascend-write-string-current-bea0d001
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-string-current-bea0d001-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,sheetjs,fastexcel-java --workload string-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-string-current-bea0d001-runs/write-string-heavy-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-string-current-bea0d001-runs/write-string-heavy-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-string-current-bea0d001-runs/write-string-heavy-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `bea0d001e77899169046d5c76d444301c269bd34`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-string-current-bea0d001`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload string-heavy`, `validationMode each`,
  `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-write-string-current-bea0d001-runs/write-string-heavy-fastest-repeat15.json
/private/tmp/ascend-write-string-current-bea0d001-runs/write-string-heavy-fastest-repeat15-scoreboard.json
```

Focused fastest comparable writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 4.121 | 15.572 | 0.612 | 90.8 MiB | 201984 |
| `fastexcel-java` | ran/lost vs Ascend | 21.711 | 36.327 | 0.194 | 792.0 MiB | 260426 |
| `excelize` | ran/lost vs Ascend | 26.365 | 29.943 | 0.051 | 23.0 MiB | 218447 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 28.206 | 35.995 | 0.075 | 23.3 MiB | 237837 |
| `sheetjs` | ran/lost vs Ascend | 39.021 | 80.256 | 0.262 | 274.0 MiB | 2016032 |

Scoreboard result:

- Focused repeat-15 fastest-writer row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The scoreboard command exits nonzero for full-profile coverage because this is
  not a full `xlsx-write-sota` run. Missing/omitted full-profile libraries,
  unsupported rows, and blocked runners are not counted as wins.

Semantic comparability: all listed rows reopened successfully, matched the
expected one-sheet and 40,000-cell shape, and passed semantic cell value
validation. Ascend, Excelize, and SheetJS matched ordered semantic cell hashes.
FastExcel Java and rust_xlsxwriter passed sorted semantic value equality but did
not match ordered semantic value hashes, so their rows are useful lower-fidelity
value-write comparisons, not byte/order-equivalent output claims. Ascend wins
median and p95 here but uses more RSS than Excelize and rust_xlsxwriter, and it
does not emit the smallest XLSX.

Humble allowed wording:

> On the generated 2000 x 20 string-heavy write row at commit `bea0d001`,
> Ascend's focused external repeat-15 run had the fastest median and p95 among
> the completed fastest comparable writers in this row. This is scoped generated
> string-write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every generated XLSX writer."
- "Ascend beats omitted, unsupported, or blocked string-heavy writers."
- "Ascend produces the smallest string-heavy XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: defer production optimization for string-heavy from this current
winning row. Continue with the next priority workflow only if it can produce a
validated optimization, comparable baseline row, or explicit claim downgrade.

## Cycle: Styles Heavy Write Fastest Comparable Row

Classification: comparable external evidence. This row is a scoped
styles-heavy write win at current `HEAD`, not a broad `xlsx-write-sota` claim.
The full profile scoreboard still fails coverage because this run intentionally
covered only one named workflow and omitted blocked/unsupported profile rows.

Workflow: generated XLSX write for styled numeric cells, 2000 rows x 20 columns.

Why it matters for release: styled generated workbook export is a user-visible
release workflow after an agent creates or edits a workbook. It exercises the
generic styled write path rather than the dense streaming path optimized for the
string-heavy row.

Public/tracked-clean input: `competitive-io` generated the `styles-heavy`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `62f45cb5`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-styles-current-62f45cb5 62f45cb5
cd /private/tmp/ascend-write-styles-current-62f45cb5
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-styles-current-62f45cb5-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,excelize,rust-xlsxwriter,fastexcel-java,pyexcelerate,pyexcelerate-cell,pyexcelerate-range,xlsxwriter,xlsxwriter-constant-memory --workload styles-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-styles-current-62f45cb5-runs/write-styles-heavy-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-styles-current-62f45cb5-runs/write-styles-heavy-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-styles-current-62f45cb5-runs/write-styles-heavy-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `62f45cb5000fedd61491c6b7dc5ec4323d3a32ff`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-styles-current-62f45cb5`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64, macOS `26.4`, kernel
  `25.4.0` on `RELEASE_ARM64_T6041`
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload styles-heavy`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-styles-current-62f45cb5-runs/write-styles-heavy-fastest-repeat15.json
/private/tmp/ascend-write-styles-current-62f45cb5-runs/write-styles-heavy-fastest-repeat15-scoreboard.json
```

Focused fastest comparable writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 28.818 | 34.161 | 0.102 | 171.4 MiB | 434417 |
| `excelize` | ran/lost vs Ascend | 41.352 | 175.258 | 0.742 | 40.5 MiB | 212204 |
| `fastexcel-java` | ran/lost vs Ascend | 52.149 | 167.356 | 0.546 | 592.0 MiB | 271770 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 119.641 | 446.356 | 0.610 | 21.2 MiB | 227730 |
| `xlsxwriter` | ran/lost vs Ascend | 250.768 | 378.144 | 0.161 | 91.3 MiB | 250499 |
| `pyexcelerate` | ran/lost vs Ascend | 259.275 | 513.785 | 0.352 | 64.7 MiB | 210468 |
| `xlsxwriter-constant-memory` | ran/lost vs Ascend | 291.538 | 418.416 | 0.183 | 48.3 MiB | 250148 |
| `pyexcelerate-range` | ran/lost vs Ascend | 327.985 | 541.478 | 0.232 | 67.7 MiB | 210468 |
| `pyexcelerate-cell` | ran/lost vs Ascend | 502.887 | 1587.896 | 0.640 | 67.1 MiB | 210468 |

Scoreboard result:

- Focused repeat-15 fastest-writer row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The scoreboard command exits nonzero for the full profile because this row is
  not full `xlsx-write-sota` coverage. Remaining coverage failures include
  other write workloads and blocked/omitted `openpyxl`, `Apache POI`, and
  `ClosedXML` styles-heavy rows. Those are not counted as wins.

Semantic comparability: all listed rows reopened successfully, matched the
expected one-sheet and 40,000-cell shape, and passed semantic cell value hash
validation. Ascend, Excelize, XlsxWriter, pyexcelerate, and XlsxWriter constant
memory also matched ordered semantic cell hashes. FastExcel Java and
rust_xlsxwriter passed sorted semantic value equality but did not match ordered
semantic value hashes, so their speed rows are useful lower-fidelity write
comparisons, not byte/order-equivalent output claims. This row validates value
semantics, not style-fidelity equivalence across writer libraries.

Humble allowed wording:

> On the generated 2000 x 20 styles-heavy write row at commit `62f45cb5`,
> Ascend's focused external repeat-15 run had the fastest median and p95 among
> the comparable completed writers in this row. This is scoped generated
> styled-write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every styles-capable writer."
- "Ascend beats OpenPyXL, Apache POI, or ClosedXML on styles-heavy writes."
- "Ascend produces the smallest styled XLSX."
- "Ascend proves style-fidelity equivalence across all compared writers."

Next action: defer production optimization for styles-heavy. Continue with the
next priority workflow only if it can produce a comparable baseline row,
validated optimization, or explicit claim downgrade.

## Cycle: Styles Heavy Write Current Fastest Comparable Row at `1908f3f5`

Classification: comparable external evidence. This refreshes the styles-heavy
write row at current `HEAD` after the latest claim-matrix commits. Ascend is the
median and p95 winner among the completed comparable writers in this row. No
production optimization is justified from a winning row.

Workflow: generated XLSX write for styled numeric cells, 2000 rows x 20 columns.

Why it matters for release: styled generated workbook export is a user-visible
agent commit/export workflow. It is also one of the named `xlsx-write-sota`
profile rows and exercises style-bearing output rather than the small dense
streaming path optimized for string-heavy writes.

Public/tracked-clean input: `competitive-io` generated the `styles-heavy`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `1908f3f5`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-styles-current-1908f3f5 1908f3f5d527
cd /private/tmp/ascend-write-styles-current-1908f3f5
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-styles-current-1908f3f5-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,excelize,rust-xlsxwriter,fastexcel-java,pyexcelerate,pyexcelerate-cell,pyexcelerate-range,xlsxwriter,xlsxwriter-constant-memory --workload styles-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-styles-current-1908f3f5-runs/write-styles-heavy-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-styles-current-1908f3f5-runs/write-styles-heavy-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-styles-current-1908f3f5-runs/write-styles-heavy-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `1908f3f5d5278fb53b45db2cf647bb5079d4e757`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-styles-current-1908f3f5`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload styles-heavy`, `validationMode each`,
  `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-write-styles-current-1908f3f5-runs/write-styles-heavy-fastest-repeat15.json
/private/tmp/ascend-write-styles-current-1908f3f5-runs/write-styles-heavy-fastest-repeat15-scoreboard.json
```

Focused fastest comparable writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 22.807 | 50.078 | 0.308 | 164.0 MiB | 434417 |
| `excelize` | ran/lost vs Ascend | 30.670 | 97.713 | 0.497 | 43.2 MiB | 212204 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 33.215 | 81.985 | 0.364 | 20.7 MiB | 227731 |
| `fastexcel-java` | ran/lost vs Ascend | 56.309 | 138.151 | 0.469 | 448.0 MiB | 271770 |
| `pyexcelerate-range` | ran/lost vs Ascend | 113.329 | 139.287 | 0.075 | 70.6 MiB | 210469 |
| `pyexcelerate` | ran/lost vs Ascend | 157.284 | 190.012 | 0.110 | 71.0 MiB | 210469 |
| `pyexcelerate-cell` | ran/lost vs Ascend | 195.004 | 315.969 | 0.213 | 72.7 MiB | 210469 |
| `xlsxwriter` | ran/lost vs Ascend | 248.385 | 392.426 | 0.267 | 100.6 MiB | 250498 |
| `xlsxwriter-constant-memory` | ran/lost vs Ascend | 251.426 | 349.041 | 0.134 | 60.0 MiB | 250149 |

Scoreboard result:

- Focused repeat-15 fastest-writer row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The scoreboard command exits nonzero for full-profile coverage because this is
  not a full `xlsx-write-sota` run. Missing/omitted full-profile libraries,
  unsupported rows, and blocked runners are not counted as wins.

Semantic comparability: all listed rows reopened successfully, matched the
expected one-sheet and 40,000-cell shape, and passed semantic cell value
validation. Ascend, Excelize, XlsxWriter, pyexcelerate, and XlsxWriter
constant-memory matched ordered semantic cell hashes. FastExcel Java and
rust_xlsxwriter passed sorted semantic value equality but did not match ordered
semantic value hashes, so their rows are useful lower-fidelity value-write
comparisons, not byte/order-equivalent output claims. This row validates value
semantics, not style-fidelity equivalence across writer libraries.

Humble allowed wording:

> On the generated 2000 x 20 styles-heavy write row at commit `1908f3f5`,
> Ascend's focused external repeat-15 run had the fastest median and p95 among
> the completed comparable writers in this row. This is scoped generated
> styled-write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every styles-capable writer."
- "Ascend beats omitted, unsupported, or blocked styles-heavy writers."
- "Ascend produces the smallest styled XLSX."
- "Ascend proves style-fidelity equivalence across all compared writers."

Next action: defer production optimization for styles-heavy from this current
winning row. Continue with the next priority workflow only if it can produce a
validated optimization, a comparable baseline row, or an explicit claim
downgrade.

## Owner-Ready Benchmark Blocker

Owner: benchmarking/external baselines.

Blocker: broad read-speed and QSS-leapfrog performance wording is downgraded.
The current full `xlsx-read-sota` profile and merged selected-sheet/metadata-only
scoreboard show no leader failures, but they still do not satisfy coverage
because blocked, unsupported, and semantic-mismatch rows remain non-wins.

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

Stop condition: do not optimize further from the measured winning rows
`dense-values`, `sparse-wide`, `styles-heavy`, `formula-heavy`, `table-heavy`,
`selected-sheet`, `metadata-only`, `warm-workflow`, and `string-heavy`. The
`feature-rich` row identified a meaningful loss and was optimized at `05656d4e`;
continue only on the next named loss, unstable tail, or memory/latency tradeoff
worth production work. The ClosedXML runner blocker was resolved by restoring
the runner in a clean worktree and is now a measured bounded-gap row. OpenPyXL
selected-sheet support and metadata-only same-lane coverage were resolved in
focused runs. The broad speed claim remains downgraded rather than promoted.

## Current Full-Profile Downgrade: XLSX Read SOTA

Classification: claim downgrade. No production optimization target is justified
from this evidence because the current full-profile and merged scoreboard report
no leader failures for completed comparable rows.

Workflow: full `xlsx-read-sota` open/inspect coverage, with focused same-lane
selected-sheet and metadata-only evidence merged back into the current profile
for claim gating.

Why it matters for release: this finishes the selected-sheet and metadata-only
timing-boundary questions while testing whether the broad speed claim can be
promoted. It cannot.

Public/tracked-clean input: `competitive-io` generated `workload all` raw OOXML
inputs from tracked benchmark code in detached commit `9ddfff91`; same-lane
selected-sheet evidence comes from clean detached commit `39163862`, and
same-lane metadata-only evidence comes from clean detached commit `fa3a13dc`.
No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-9ddfff91 9ddfff91
cd /private/tmp/ascend-perf-hillclimb-9ddfff91
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload all --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-all.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-all.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-all-scoreboard.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-merged-selected-metadata.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-merged-selected-metadata-scoreboard.json
```

Environment:

- Commit: `9ddfff91efc8f0f95edf36f44b78f5313480ad11`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-9ddfff91`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Node: `24.3.0`
- Runtime profile: `category read`, `competitor all`, `workload all`, `readSource raw-ooxml`, `validationMode each`, `repeat 5`, `warmup 1`, `sourceMode full`.

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-all.json
/private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-all-scoreboard.json
/private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-merged-selected-metadata.json
/private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-merged-selected-metadata-scoreboard.json
```

Scoreboard result:

- Current full-profile scoreboard: `leaderFailures: []`,
  `profileLeaderFailures: []`, `coverageFailures: 12`.
- Merged selected-sheet/metadata-only scoreboard: `leaderFailures: []`,
  `profileLeaderFailures: []`, `coverageFailures: 10`, `coverageGaps: 8`.
- The merged scoreboard removes the selected-sheet and metadata-only
  `missing-comparable` failures. The remaining `coverageFailures` are ClosedXML
  missing/error rows for required value-read workloads and SheetJS/Calamine
  semantic mismatches for `feature-rich` rich metadata.

Representative Ascend medians from the current full-profile run:

| Workload | Representative Ascend row | Median ms | P95 ms | Peak RSS | Decision |
| --- | --- | ---: | ---: | ---: | --- |
| `dense-values` | `ascend-readxlsx-raw-values-operation-bytes` | 3.166 | 3.287 | 91.1 MiB | no optimization target |
| `sparse-wide` | `ascend-readxlsx-row-stream-bytes` | 8.223 | 9.052 | 100.3 MiB | no optimization target |
| `string-heavy` | `ascend-readxlsx-row-stream-bytes` | 8.412 | 8.451 | 111.7 MiB | no optimization target |
| `styles-heavy` | `ascend` | 4.849 | 5.515 | 148.2 MiB | no optimization target |
| `formula-heavy` | `ascend-readxlsx-raw-values-operation-bytes` | 5.338 | 6.074 | 100.7 MiB | no optimization target |
| `table-heavy` | `ascend-readxlsx-row-stream-bytes` | 5.608 | 5.743 | 98.8 MiB | no optimization target |
| `feature-rich` | `ascend` | 4.955 | 5.161 | 157.2 MiB | claim blocked by competitor semantic mismatch |
| `warm-workflow` | `ascend` | 2.870 | 3.545 | 183.7 MiB | no optimization target |

Humble allowed wording:

> In a clean detached current full-profile run at `9ddfff91`, Ascend reported no leader failures for completed comparable `xlsx-read-sota` rows. Merged same-lane selected-sheet and metadata-only evidence removes those comparability failures, but the broad speed claim remains downgraded because ClosedXML coverage, feature-rich semantic mismatches, and unsupported selected-sheet/metadata-only competitors remain non-wins.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend is SOTA for XLSX read."
- "Ascend beats ClosedXML, fastxlsx, unsupported runners, or semantic-mismatch rows."
- Any wording that converts `leaderFailures: []` into a full-profile coverage pass.

Next action: stop production optimization from this evidence. If the performance
loop continues, it must be blocker work: ClosedXML coverage policy, feature-rich
SheetJS/Calamine semantic support or not-comparable policy, unsupported
selected-sheet/metadata-only competitor policy, or FastXLSX environment coverage.

## Full Current-Commit Gate: XLSX Read SOTA

Classification: blocked/defer. No production optimization is justified from
this gate because the completed comparable head-to-head rows produced no leader
failures; broad wording is still blocked by explicit coverage and semantic
comparability gaps.

Workflow: full `xlsx-read-sota` open/inspect coverage across the tracked read
workloads.

Why it matters for release: this is the closest current evidence to a broad
release-read claim. It checks whether the per-workload wins survive a single
current-commit run, and it prevents treating failed, unsupported, or
semantically mismatched external runners as Ascend wins.

Public/tracked-clean input: `competitive-io` generated `workload all` raw OOXML
inputs from tracked benchmark code in detached commit `4b1c1734`. The profile
included the required release rows `dense-values`, `sparse-wide`,
`string-heavy`, `styles-heavy`, `formula-heavy`, `table-heavy`, `feature-rich`,
`selected-sheet`, `metadata-only`, and `warm-workflow`; the harness also emitted
non-profile generated read workloads. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-4b1c1734 4b1c1734
cd /private/tmp/ascend-perf-hillclimb-4b1c1734
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload all --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-4b1c1734-runs/xlsx-read-sota-all.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-4b1c1734-runs/xlsx-read-sota-all.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-4b1c1734-runs/xlsx-read-sota-all-scoreboard.json
```

Environment:

- Commit: `4b1c1734b95ee96da0078b60f5e439768303e04e`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-4b1c1734`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Node: `24.3.0`
- Runtime profile: `category read`, `competitor all`, `workload all`, `readSource raw-ooxml`, `validationMode each`, `repeat 5`, `warmup 1`, `sourceMode full`.

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-4b1c1734-runs/xlsx-read-sota-all.json
/private/tmp/ascend-perf-hillclimb-4b1c1734-runs/xlsx-read-sota-all-scoreboard.json
```

Scoreboard result:

- `leaderFailures: []`
- `profileLeaderFailures: []`
- Coverage failures remain:
  `ClosedXML` missing for `dense-values`, `sparse-wide`, `string-heavy`,
  `styles-heavy`, `formula-heavy`, `table-heavy`, `feature-rich`, and
  `warm-workflow`; `SheetJS` and `Calamine` ineligible for `feature-rich`
  rich-metadata read because their correctness status is `semantic-mismatch`;
  and `metadata-only` reports `missing-comparable` for required competitors
  `Ascend`, `SheetJS`, and `openpyxl`.
- Coverage gaps remain: `selected-sheet` is `unsupported-operation` for
  `ExcelJS`, `Apache POI`, and `ClosedXML`;
  `metadata-only` is `unsupported-operation` for `ExcelJS`, `Calamine`,
  `Apache POI`, and `ClosedXML`.

Competitor status:

| Competitor/gap | Status | Semantic comparability |
| --- | --- | --- |
| Completed comparable profile rows | ran/won | No `leaderFailures` or `profileLeaderFailures` were reported by the current full-profile scoreboard. This supports scoped wording only for completed comparable rows. |
| ClosedXML | blocked | Missing across most required read-value rows because the runner remains blocked. Not counted as an Ascend win. |
| SheetJS and Calamine on `feature-rich` rich metadata | not comparable | The scoreboard marks these rows ineligible with `correctnessStatus=semantic-mismatch`. Not counted as Ascend wins. |
| `selected-sheet` unsupported competitors | not comparable | ExcelJS, Apache POI, and ClosedXML do not provide the same selected-sheet operation in this profile. Not counted as wins. |
| `metadata-only` unsupported competitors | not comparable | ExcelJS, Calamine, Apache POI, and ClosedXML are unsupported for the metadata-only operation; the full scoreboard also reports a metadata-only `missing-comparable` grouping gap for Ascend, SheetJS, and openpyxl. |
| fastxlsx | runner unavailable | Not part of a completed comparable current full-profile row; missing dependency rows remain non-wins. |

Humble allowed wording:

> In a clean detached current-commit full `xlsx-read-sota` run at `4b1c1734`, the scoreboard reported no leader failures for completed comparable profile rows. A broad XLSX-read speed claim is still not promotable because ClosedXML coverage, feature-rich semantic mismatches, selected-sheet unsupported operations, and metadata-only comparability gaps remain explicit blockers.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend is SOTA for XLSX read."
- "Ascend beats ClosedXML, fastxlsx, unsupported runners, or semantic-mismatch rows."
- Any wording that converts `leaderFailures: []` into a full-profile coverage pass.

Next action: attack the highest-impact blocker, starting with the ClosedXML
runner because it blocks the most required head-to-head read rows. If the runner
cannot be made comparable in the local release environment, record a tighter
blocked decision and keep the broad claim downgraded.

## Cycle: ClosedXML Focused Head-to-Head Read

Classification: defer. ClosedXML is no longer a runner blocker for the required
value-read rows, but no production optimization is justified because the focused
clean head-to-head run shows Ascend ahead on the comparable workflows.

Workflow: XLSX open/inspect value read against ClosedXML for the required
`xlsx-read-sota` read rows.

Why it matters for release: ClosedXML was the largest remaining missing
competitor in the full current-commit gate. Converting it from `blocked` to
measured rows strengthens the external claim matrix without counting a failed
runner as a win.

Public/tracked-clean input: `competitive-io` generated `workload all` raw OOXML
inputs from tracked benchmark code in detached commit `2f5c1761`. Required
profile rows covered here are `dense-values`, `sparse-wide`, `string-heavy`,
`styles-heavy`, `formula-heavy`, `table-heavy`, `feature-rich`, and
`warm-workflow`. `selected-sheet` and `metadata-only` remain unsupported for
ClosedXML in the profile and are not counted as wins.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-2f5c1761 2f5c1761
cd /private/tmp/ascend-perf-hillclimb-2f5c1761
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin DOTNET_CLI_HOME=/private/tmp/ascend-dotnet-home NUGET_PACKAGES=/private/tmp/ascend-nuget-packages dotnet build fixtures/benchmarks/runners/closedxml/ClosedXmlRunner.csproj --configuration Release -v minimal
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor external --libraries closedxml --workload all --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-2f5c1761-runs/closedxml-read-all-after-restore.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor external --libraries ascend-readxlsx-raw-values-operation-path,ascend-readxlsx-raw-values-operation-bytes,ascend-readxlsx-values-rich-metadata-bytes,closedxml --workload all --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-2f5c1761-runs/ascend-closedxml-head-to-head-all.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-2f5c1761-runs/ascend-closedxml-head-to-head-all.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-2f5c1761-runs/ascend-closedxml-head-to-head-all-scoreboard.json
```

Environment:

- Commit: `2f5c17617ae2c41cac84558edebe3b3174c30a09`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-2f5c1761`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Node: `24.3.0`
- .NET SDK: `8.0.125`
- ClosedXML runner version: `0.105.0.0`
- Runtime profile: `category read`, `competitor external`, `workload all`, `readSource raw-ooxml`, `validationMode each`, `repeat 5`, `warmup 1`, `sourceMode full`.

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-2f5c1761-runs/closedxml-read-all-after-restore.json
/private/tmp/ascend-perf-hillclimb-2f5c1761-runs/ascend-closedxml-head-to-head-all.json
/private/tmp/ascend-perf-hillclimb-2f5c1761-runs/ascend-closedxml-head-to-head-all-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. The
Ascend column uses the fastest eligible Ascend row for the workflow from the
same focused harness invocation. These rows support bounded ClosedXML
comparisons only; they do not convert the full profile into a coverage pass.

| Workflow | Ascend row | Ascend median / p95 / CV / RSS | ClosedXML median / p95 / CV / RSS | Status | Semantic comparability |
| --- | --- | ---: | ---: | --- | --- |
| `dense-values` | `ascend-readxlsx-raw-values-operation-bytes` | 9.394 ms / 11.905 ms / 0.165 / 85.6 MiB | 369.344 ms / 884.345 ms / 0.514 / 113.4 MiB | ran/won | Value-read assertions pass for both. ClosedXML materializes a workbook; use only as a bounded gap for open/inspect value reads. |
| `sparse-wide` | `ascend-readxlsx-raw-values-operation-bytes` | 18.495 ms / 50.346 ms / 0.628 / 119.6 MiB | 507.483 ms / 683.386 ms / 0.261 / 144.4 MiB | ran/won | Value-read assertions pass for both; Ascend tail is noisy. |
| `string-heavy` | `ascend-readxlsx-raw-values-operation-bytes` | 26.880 ms / 52.719 ms / 0.355 / 107.2 MiB | 320.998 ms / 1353.840 ms / 0.911 / 123.6 MiB | ran/won | Value-read assertions pass for both; ClosedXML tail is very noisy. |
| `styles-heavy` | `ascend-readxlsx-raw-values-operation-path` | 13.035 ms / 14.260 ms / 0.057 / 101.2 MiB | 322.689 ms / 560.705 ms / 0.306 / 123.0 MiB | ran/won | Value-read assertions pass for both; this does not prove style fidelity or preservation. |
| `formula-heavy` | `ascend-readxlsx-raw-values-operation-path` | 10.298 ms / 10.660 ms / 0.029 / 96.5 MiB | 358.916 ms / 513.109 ms / 0.234 / 132.7 MiB | ran/won | Value-read assertions pass for both; this does not prove formula calculation or preservation. |
| `table-heavy` | `ascend-readxlsx-raw-values-operation-bytes` | 19.329 ms / 32.379 ms / 0.320 / 114.0 MiB | 307.502 ms / 413.116 ms / 0.210 / 125.4 MiB | ran/won | Value-read assertions pass for both; this does not prove table preservation. |
| `feature-rich` | `ascend-readxlsx-values-rich-metadata-bytes` | 50.600 ms / 52.510 ms / 0.036 / 169.3 MiB | 194.255 ms / 610.081 ms / 0.618 / 142.5 MiB | ran/won | Rich-metadata assertions pass for both; ClosedXML has lower RSS but a much slower and noisier tail. |
| `warm-workflow` | `ascend-readxlsx-raw-values-operation-bytes` | 3.544 ms / 3.623 ms / 0.028 / 90.8 MiB | 249.849 ms / 318.092 ms / 0.275 / 120.8 MiB | ran/won | Warm value-read assertions pass for both. |
| `selected-sheet` | n/a | n/a | n/a | not comparable | ClosedXML remains `unsupported-operation` for selected-sheet read in the profile. |
| `metadata-only` | n/a | n/a | n/a | not comparable | ClosedXML remains `unsupported-operation` for metadata-only read in the profile. |

Scoreboard result for the focused run:

- `leaderFailures: []`
- `profileLeaderFailures: []`
- The scoreboard still fails full profile coverage because this focused run
  intentionally omits SheetJS, ExcelJS, openpyxl, Calamine, and Apache POI rows,
  and because selected-sheet and metadata-only unsupported-operation gaps remain.

Humble allowed wording:

> In a clean focused `xlsx-read-sota` head-to-head run at `2f5c1761`, ClosedXML `0.105.0.0` successfully ran the required comparable value-read rows and was slower than Ascend's measured value-read path on those workflows. This is bounded ClosedXML evidence only, not a full-profile XLSX-read speed claim.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend is SOTA for XLSX read."
- "Ascend beats ClosedXML for selected-sheet or metadata-only reads."
- Any wording that hides the different timing models, noisy tails, or remaining full-profile coverage gaps.

Next action: defer production optimization. The next highest-impact gaps are no
longer ClosedXML value-read coverage; they are selected-sheet/metadata-only
unsupported-operation gaps, feature-rich semantic mismatches for SheetJS and
Calamine, and the unavailable fastxlsx runner.

## Cycle: Selected-Sheet OpenPyXL Head-to-Head Read

Classification: defer. The openpyxl selected-sheet gap moved from
`unsupported-operation` to a passing measured row, but no production
optimization is justified because Ascend still leads the focused comparison and
the full profile needs timing-lane alignment before promotion.

Workflow: XLSX selected-sheet open/inspect value read for the `Data` sheet while
the generated workbook also contains `Summary` and `Archive`.

Why it matters for release: selected-sheet read is a user-visible agent workflow
for inspecting one sheet of a multi-sheet workbook without hydrating unrelated
sheets. openpyxl was previously listed as unsupported for this profile gap; this
cycle adds a real external baseline instead of counting the absence as a win.

Public/tracked-clean input: `competitive-io` generated the `selected-sheet`
raw OOXML workload from detached commit `57c5a242`, 2000 rows x 20 columns,
40,000 `Data` cells plus auxiliary `Summary` and `Archive` sheets, 114,747 input
bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-57c5a242 57c5a242
cd /private/tmp/ascend-perf-hillclimb-57c5a242
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --libraries ascend,sheetjs,openpyxl --workload selected-sheet --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-57c5a242-runs/selected-sheet-openpyxl-head-to-head.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-57c5a242-runs/selected-sheet-openpyxl-head-to-head.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-57c5a242-runs/selected-sheet-openpyxl-head-to-head-scoreboard.json
```

Environment:

- Commit: `57c5a2420d9d19a3f4f2138bf6644303450b01a1`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-57c5a242`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Node: `24.3.0`
- Runtime profile: `category read`, `workload selected-sheet`, `readSource raw-ooxml`, `validationMode each`, `repeat 5`, `warmup 1`.

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-57c5a242-runs/selected-sheet-openpyxl-head-to-head.json
/private/tmp/ascend-perf-hillclimb-57c5a242-runs/selected-sheet-openpyxl-head-to-head-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. All
three rows assert `selectedSheetRead: true`, `sourceSheetCount: 3`,
`loadedSheetCount: 1`, `loadedSheetNames: Data`, `hasAllSheets: false`, and
`selectedSheetMatches: true`.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend` | 37.985 | 82.642 | 0.726 | 185.6 MiB | In-process selected-sheet value read. Correct semantics, but tail is very noisy in this clean run. |
| SheetJS | ran/lost | `sheetjs` | 66.446 | 82.170 | 0.119 | 237.6 MiB | In-process selected-sheet parse of `Data`; 1.75x slower than Ascend by median. |
| openpyxl | ran/lost | `openpyxl` | 372.944 | 405.825 | 0.076 | 87.4 MiB | External-process selected-sheet assertions over `Data`; lower memory than Ascend but 9.82x slower by median. |
| ExcelJS | not comparable | n/a | n/a | n/a | n/a | n/a | Still skipped: selected-sheet read is unsupported by this harness without full workbook hydration. |
| Calamine | not comparable | n/a | n/a | n/a | n/a | n/a | Still a profile `unsupported-operation` gap. |
| Apache POI | not comparable | n/a | n/a | n/a | n/a | n/a | Still a profile `unsupported-operation` gap. |
| ClosedXML | not comparable | n/a | n/a | n/a | n/a | n/a | Still a profile `unsupported-operation` gap. |

Scoreboard result for the focused run:

- `leaderFailures: []`
- `profileLeaderFailures: []`
- The selected-sheet profile row now requires `openpyxl`, but the focused
  scoreboard still reports `missing-comparable` for
  `requiredCompetitors=Ascend,SheetJS,openpyxl` because the current evidence
  spans in-process and external-process timing lanes.
- The openpyxl selected-sheet `unsupported-operation` coverage gap is removed
  from the current scorer; ExcelJS, Calamine, Apache POI, and ClosedXML remain
  unsupported-operation gaps.

Humble allowed wording:

> In a clean focused selected-sheet run at `57c5a242`, Ascend, SheetJS, and openpyxl all loaded only the `Data` sheet from a three-sheet generated workbook. Ascend had the fastest median in that focused comparison, but the full profile still needs same-lane evidence before broad selected-sheet wording is promotable.

Forbidden wording:

- "Ascend has a full selected-sheet SOTA claim."
- "Ascend beats ExcelJS, Calamine, Apache POI, or ClosedXML on selected-sheet reads."
- Any wording that hides Ascend's noisy tail or the in-process/external-process lane split.

Next action: defer production optimization. Add external-process selected-sheet
lanes for Ascend and SheetJS, or record the lane split as a permanent
not-comparable boundary if those runners cannot expose equivalent semantics.

## Cycle: Selected-Sheet Same-Lane External Read

Classification: defer. The previous selected-sheet baseline was cooked by a
mixed in-process/external timing lane and by an Ascend external assertion bug
that reported unloaded source-sheet placeholders as loaded sheets. Both issues
are fixed in this cycle. No production optimization is justified because Ascend
now has a clean same-lane external-process win.

Workflow: XLSX selected-sheet open/inspect value read for the `Data` sheet from
a three-sheet workbook.

Why it matters for release: this turns the selected-sheet row from a bounded but
not-comparable baseline into a comparable external-process head-to-head row for
Ascend, SheetJS, and openpyxl.

Public/tracked-clean input: `competitive-io` generated the `selected-sheet`
raw OOXML workload from detached commit `39163862`, 2000 rows x 20 columns,
40,000 `Data` cells plus auxiliary `Summary` and `Archive` sheets, 114,747 input
bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-39163862 39163862
cd /private/tmp/ascend-perf-hillclimb-39163862
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-values,sheetjs,openpyxl --workload selected-sheet --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/selected-sheet-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-39163862-runs/selected-sheet-same-lane.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-39163862-runs/selected-sheet-same-lane.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-39163862-runs/selected-sheet-same-lane-scoreboard.json
```

Environment:

- Commit: `3916386295e83b234dba10bdb7f007a9f5d52704`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-39163862`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Node: `24.3.0`
- Runtime profile: `category read`, `executionScope external-process`, `workload selected-sheet`, `readSource raw-ooxml`, `validationMode each`, `repeat 5`, `warmup 1`.

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-39163862-runs/selected-sheet-same-lane.json
/private/tmp/ascend-perf-hillclimb-39163862-runs/selected-sheet-same-lane-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. All
three rows share `external-internal-operation-timing:selected-sheet` and assert
`selectedSheetRead: true`, `sourceSheetCount: 3`, `loadedSheetCount: 1`,
`loadedSheetNames: Data`, `hasAllSheets: false`, and `selectedSheetMatches:
true`.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend-external-values` | 22.066 | 25.622 | 0.085 | 154.4 MiB | External-process selected-sheet value read from file path; same timing lane as SheetJS and openpyxl. |
| SheetJS | ran/lost | `sheetjs` | 27.981 | 31.732 | 0.068 | 258.6 MiB | External-process selected-sheet parse of `Data`; 1.27x slower than Ascend by median. |
| openpyxl | ran/lost | `openpyxl` | 205.474 | 218.842 | 0.059 | 87.2 MiB | External-process selected-sheet assertions over `Data`; lower memory than Ascend but 9.31x slower by median. |
| ExcelJS | not comparable | n/a | n/a | n/a | n/a | n/a | Still skipped: selected-sheet read is unsupported by this harness without full workbook hydration. |
| Calamine | not comparable | n/a | n/a | n/a | n/a | n/a | Still a profile `unsupported-operation` gap. |
| Apache POI | not comparable | n/a | n/a | n/a | n/a | n/a | Still a profile `unsupported-operation` gap. |
| ClosedXML | not comparable | n/a | n/a | n/a | n/a | n/a | Still a profile `unsupported-operation` gap. |

Scoreboard result for the focused same-lane run:

- `leaderFailures: []`
- `profileLeaderFailures: []`
- No selected-sheet `coverageFailures` remain for Ascend, SheetJS, or openpyxl.
- Selected-sheet `coverageGaps` remain for ExcelJS, Apache POI, and ClosedXML
  because their selected-sheet operation is unsupported in the current profile.

Humble allowed wording:

> On the generated `selected-sheet` raw OOXML workload at commit `39163862`, Ascend's external-process selected-sheet value read was faster by median than the same-lane SheetJS and openpyxl rows that successfully ran. ExcelJS, Apache POI, and ClosedXML remain unsupported-operation gaps, and python-calamine has only current-worktree runner proof, so this is scoped selected-sheet evidence, not a broad XLSX-read claim.

Forbidden wording:

- "Ascend has a full selected-sheet SOTA claim across every library."
- "Ascend beats ExcelJS, Calamine, Apache POI, or ClosedXML on selected-sheet reads."
- Any wording that treats unsupported selected-sheet competitors as wins.

Next action: defer production optimization. The next highest-impact blockers are
metadata-only same-lane coverage, feature-rich semantic mismatches for SheetJS
and Calamine, and the unavailable fastxlsx runner.

## Cycle: Selected-Sheet Current Same-Lane Read

Classification: comparable external evidence plus timing-boundary downgrade.
The current clean same-lane row confirms a `readXlsx` selected-sheet win against
SheetJS, openpyxl, and python-calamine. It also exposes a boundary: the
`ascend-external-values` SDK runner is not used for speed wording here because
it times assertion/materialization work that the external open-only rows do not.
No production optimization is justified from this evidence.

Workflow: XLSX selected-sheet open/inspect value read for the `Data` sheet from
a generated three-sheet workbook.

Why it matters for release: selected-sheet read is the first open/inspect step
for agents that only need one worksheet from a larger workbook. This row closes
the current selected-sheet timing-boundary rerun without counting unsupported
or mixed-timing rows as wins.

Public/tracked-clean input: `competitive-io` generated the `selected-sheet`
raw OOXML workload from tracked benchmark code in a clean detached worktree at
commit `27af69d4`, 2000 rows x 20 columns, 40,000 `Data` cells plus auxiliary
`Summary` and `Archive` sheets, 114,747 input bytes. No private corpus or local
research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-selected-current-27af69d4 27af69d4
cd /private/tmp/ascend-selected-current-27af69d4
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-selected-current-27af69d4-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-values,ascend-readxlsx-selected-values,sheetjs,openpyxl,python-calamine --workload selected-sheet --read-source raw-ooxml --repeat 15 --warmup 3 --validation-mode each --runner-manifest fixtures/benchmarks/runners/selected-sheet-readers.manifest.json > /private/tmp/ascend-selected-current-27af69d4-runs/selected-sheet-same-lane-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-selected-current-27af69d4-runs/selected-sheet-same-lane-repeat15.json --json --metric medianMs --require-profile xlsx-read-sota --assert-profile-leader ascend > /private/tmp/ascend-selected-current-27af69d4-runs/selected-sheet-same-lane-repeat15-scoreboard.json
```

Environment:

- Commit: `27af69d411dec7007c32d29e0727ce02d4662e84`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-selected-current-27af69d4`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Python: `3.13.3`
- Platform: Darwin arm64, macOS `26.4`, kernel
  `25.4.0` on `RELEASE_ARM64_T6041`
- Runtime profile: `category read`, `executionScope external-process`,
  `workload selected-sheet`, `readSource raw-ooxml`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-selected-current-27af69d4-runs/selected-sheet-same-lane-repeat15.json
/private/tmp/ascend-selected-current-27af69d4-runs/selected-sheet-same-lane-repeat15-scoreboard.json
```

Focused same-lane selected-sheet row, repeat 15 after 3 warmups:

| Runner | Status vs `readXlsx` | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `ascend-readxlsx-selected-values` | ran/won | 6.579 | 8.909 | 0.173 | 95.3 MiB | `readXlsx` selected-sheet open-only row. Loads only `Data`; shape and value hashes match. |
| `python-calamine` | ran/lost vs `readXlsx` | 25.791 | 72.888 | 0.413 | 44.0 MiB | Same selected-sheet value semantics; lower RSS but noisier and 3.92x slower by median. |
| `sheetjs` | ran/lost vs `readXlsx` | 58.808 | 80.457 | 0.128 | 254.9 MiB | Same selected-sheet value semantics; ordered value hash matches. |
| `openpyxl` | ran/lost vs `readXlsx` | 413.117 | 689.932 | 0.226 | 96.8 MiB | Same selected-sheet value semantics; much slower by median. |
| `ascend-external-values` | timing-boundary, not used for speed wording | 49.792 | 68.437 | 0.150 | 156.5 MiB | SDK `Ascend.open` row loads only `Data`, but the runner times assertion/materialization work; do not compare it to open-only rows. |

Scoreboard result:

- Focused selected-sheet row winner: `ascend-readxlsx-selected-values`.
- `leaderFailures: []` and `profileLeaderFailures: []`.
- The scoreboard command exits nonzero for the full profile because this row is
  not full `xlsx-read-sota` coverage.
- Remaining selected-sheet `coverageGaps` include ExcelJS, Apache POI, and
  ClosedXML unsupported-operation rows. The scorer also still reports a
  `Calamine` unsupported-operation profile gap even though `python-calamine`
  ran in this focused row, so broad Calamine wording remains downgraded until
  profile policy and runner naming agree.

Semantic comparability: all listed rows assert `selectedSheetRead: true`,
`sourceSheetCount: 3`, `loadedSheetCount: 1`, `loadedSheetNames: Data`,
`hasAllSheets: false`, `selectedSheetMatches: true`, `cellCountMatches: true`,
and `semanticCellValuesHashMatches: true`. The `readXlsx`, python-calamine,
openpyxl, and SDK Ascend rows do not match ordered semantic value hashes, so
ordered output claims are forbidden. This is selected-sheet value-read evidence,
not style, formula calculation, package preservation, or edit-safety evidence.

Humble allowed wording:

> On the generated `selected-sheet` raw OOXML workload at commit `27af69d4`,
> Ascend's `readXlsx` selected-sheet open row was faster by median and p95 than
> the same-lane SheetJS, openpyxl, and python-calamine rows that successfully
> ran. This is scoped selected-sheet value-read evidence, not a broad
> `xlsx-read-sota` claim.

Forbidden wording:

- "Ascend has a full selected-sheet SOTA claim."
- "Ascend beats ExcelJS, Apache POI, or ClosedXML on selected-sheet reads."
- "Ascend beats Calamine in the full profile" while the scorer still reports a
  Calamine selected-sheet coverage gap.
- "Ascend's SDK `Ascend.open` selected-sheet path beats python-calamine."
- Any wording that treats unsupported competitors, mixed timing boundaries, or
  ordered-hash mismatches as wins.

Next action: defer production optimization for low-level `readXlsx`
selected-sheet. If release wording needs the SDK `Ascend.open` selected-sheet
workflow, first produce a same-timing SDK open-only row or classify the current
SDK runner as a benchmark timing bug before optimizing production code.

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

## Cycle: Formula-Heavy Value Read

Classification: defer. No production optimization is justified from this cycle.

Workflow: XLSX open/inspect value read for a formula-heavy worksheet.

Why it matters for release: this is a required `xlsx-read-sota` workload and a common agent inspection path for model-like workbooks where formula cells are present but the first release workflow only needs cached/value inspection before planning.

Public/tracked-clean input: `competitive-io` generated `formula-heavy` workbook using tracked benchmark code from detached commit `e1c69a32`, `raw-ooxml` source, 2000 rows x 20 columns, 40,000 logical cells, 40,000 populated cells, 246,241 input bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-e1c69a32 e1c69a32
cd /private/tmp/ascend-perf-hillclimb-e1c69a32
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload formula-heavy --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-e1c69a32-runs/formula-heavy-read-values.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-e1c69a32-runs/formula-heavy-read-values.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-e1c69a32-runs/formula-heavy-scoreboard.json
```

Environment:

- Commit: `e1c69a32`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-e1c69a32`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Python: `3.13.3`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-e1c69a32-runs/formula-heavy-read-values.json
/private/tmp/ascend-perf-hillclimb-e1c69a32-runs/formula-heavy-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. Rows marked `blocked`, `runner unavailable`, or `not comparable` are non-ranking status rows for claim wording and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 5.587 | 5.689 | 0.013 | 101.9 MiB | File-path value read with post-operation assertions. Comparable for cached/value inspection only; this does not prove formula calculation or formula preservation. |
| Ascend row stream | not comparable | `ascend-readxlsx-row-stream-bytes` | 7.019 | 7.123 | 0.036 | 103.5 MiB | Row-stream over bytes; useful diagnostic path but not the default file-path comparison. |
| fastexcel Python | ran/lost | `fastexcel` | 9.930 | 10.312 | 0.018 | 63.9 MiB | Value read, operation timing. Lower memory; bounded gap versus Ascend file-path is 1.78x slower by median. |
| FastExcel Java | ran/lost | `fastexcel-java` | 13.139 | 32.818 | 0.520 | 344.0 MiB heap | Streaming value read. Tail is noisy; bounded gap is 2.35x slower by median. |
| Polars calamine | ran/lost | `polars-calamine` | 13.702 | 15.942 | 0.119 | 103.6 MiB | Value read through Polars/calamine. Bounded gap is 2.45x slower by median. |
| python-calamine | ran/lost | `python-calamine` | 20.943 | 21.809 | 0.020 | 42.0 MiB | File-path materialized value read. Lower memory; 3.75x slower by median. |
| rust-calamine | ran/lost | `rust-calamine` | 37.957 | 38.907 | 0.020 | 12.2 MiB | File-path materialized value read. Much lower memory; 6.79x slower by median. |
| SheetJS | ran/lost | `sheetjs` | 45.091 | 46.836 | 0.040 | 291.5 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| ExcelJS | ran/lost | `exceljs` | 59.203 | 63.610 | 0.085 | 343.9 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| Apache POI | ran/lost | `apache-poi` | 89.125 | 144.209 | 0.290 | 352.0 MiB heap | Materialized workbook value read. Tail is moderately noisy; 15.95x slower by median. |
| Polars xlsx2csv | ran/lost | `polars-xlsx2csv` | 97.886 | 125.630 | 0.133 | 75.3 MiB | Value read through xlsx2csv path. Tail is moderately noisy; 17.52x slower by median. |
| Polars openpyxl | ran/lost | `polars-openpyxl` | 173.149 | 183.207 | 0.072 | 118.9 MiB | Value read through openpyxl engine. 30.99x slower by median. |
| openpyxl read-only values | ran/lost | `openpyxl-read-only-values` | 188.958 | 191.213 | 0.017 | 64.8 MiB | Streaming data-only value read. Lower memory; 33.82x slower by median. |
| pyopenxlsx | ran/lost | `pyopenxlsx` | 244.477 | 248.577 | 0.015 | 84.2 MiB | Cell materialization. 43.76x slower by median. |
| Excelize | ran/lost | `excelize` | 303.583 | 342.048 | 0.295 | 59.6 MiB | File-path materialized value read. Lower memory; 54.34x slower by median. |
| ClosedXML | blocked | `closedxml` | n/a | n/a | n/a | n/a | Not counted. The .NET runner failed with `CSSM_ModuleLoad(): One or more parameters passed to a function were not valid.` |
| fastxlsx | runner unavailable | `fastxlsx` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `fastxlsx` is not installed in the clean Python environment. |

In-process JavaScript reference:

| Library | Status | Median ms | P95 ms | CV | Peak RSS |
| --- | --- | ---: | ---: | ---: | ---: |
| Ascend in-process | ran/won | 5.389 | 5.905 | 0.056 | 196.9 MiB |
| SheetJS `xlsx@0.18.5` | ran/lost | 45.091 | 46.836 | 0.040 | 291.5 MiB |
| ExcelJS `4.4.0` | ran/lost | 59.203 | 63.610 | 0.085 | 343.9 MiB |

Coverage gate result: failed, as expected for a partial profile. The formula-heavy row lacks ClosedXML and fastxlsx coverage. The profile remains missing `string-heavy` from current commit, `table-heavy`, `feature-rich`, `selected-sheet`, `metadata-only`, and `warm-workflow` coverage.

Humble allowed wording:

> On the generated `formula-heavy` 2000x20 raw OOXML workload at commit `e1c69a32`, Ascend's file-path value-read row had the fastest eligible median and p95 among completed comparable external runners. The full `xlsx-read-sota` profile remains incomplete, and this row proves value inspection only, not formula calculation.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend has a clean formula-heavy read speed win across all runners."
- "Ascend beats ClosedXML or fastxlsx" from this run.
- "Ascend proves formula calculation, recalc, or formula preservation" from this value-read run.
- Any wording that hides the incomplete full profile or blocked/unavailable runners.

Next action: defer production optimization. Continue profile expansion with `table-heavy`; keep ClosedXML and fastxlsx as blockers unless their runners are fixed.

## Cycle: Table-Heavy Value Read

Classification: defer. No production optimization is justified from this cycle.

Workflow: XLSX open/inspect value read for a table-heavy worksheet.

Why it matters for release: this is a required `xlsx-read-sota` workload and a common operational workbook shape where structured table metadata surrounds ordinary values.

Public/tracked-clean input: `competitive-io` generated `table-heavy` workbook using tracked benchmark code from detached commit `65520519`, `raw-ooxml` source, 2000 rows x 20 columns, 40,000 logical cells, 40,000 populated cells, 151,994 input bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-65520519 65520519
cd /private/tmp/ascend-perf-hillclimb-65520519
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload table-heavy --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-65520519-runs/table-heavy-read-values.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-65520519-runs/table-heavy-read-values.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-65520519-runs/table-heavy-scoreboard.json
```

Environment:

- Commit: `65520519`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-65520519`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Python: `3.13.3`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-65520519-runs/table-heavy-read-values.json
/private/tmp/ascend-perf-hillclimb-65520519-runs/table-heavy-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. Rows marked `blocked`, `runner unavailable`, or `not comparable` are non-ranking status rows for claim wording and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 8.822 | 9.551 | 0.052 | 115.2 MiB | File-path value read with post-operation assertions. Comparable for value-read inspection only; this does not prove table metadata fidelity or preservation. |
| Ascend row stream | not comparable | `ascend-readxlsx-row-stream-bytes` | 5.405 | 5.610 | 0.023 | 99.6 MiB | Row-stream over bytes; faster diagnostic path but not the default file-path comparison. |
| FastExcel Java | ran/lost | `fastexcel-java` | 12.509 | 28.844 | 0.443 | 112.0 MiB heap | Streaming value read. Tail is noisy; bounded gap versus Ascend file-path is 1.42x slower by median. |
| python-calamine | ran/lost | `python-calamine` | 20.999 | 26.525 | 0.133 | 44.1 MiB | File-path materialized value read. Lower memory; 2.38x slower by median. |
| SheetJS | ran/lost | `sheetjs` | 37.050 | 43.049 | 0.084 | 312.2 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| rust-calamine | ran/lost | `rust-calamine` | 42.954 | 48.639 | 0.074 | 14.4 MiB | File-path materialized value read. Much lower memory; 4.87x slower by median. |
| ExcelJS | ran/lost | `exceljs` | 53.997 | 61.667 | 0.069 | 324.0 MiB | In-process value read. Compare only with Ascend in-process row, not external-process lanes. |
| Apache POI | ran/lost | `apache-poi` | 132.726 | 247.560 | 0.421 | 352.0 MiB heap | Materialized workbook value read. Tail is noisy; 15.04x slower by median. |
| Excelize | ran/lost | `excelize` | 153.556 | 190.546 | 0.104 | 54.6 MiB | File-path materialized value read. Lower memory; 17.41x slower by median. |
| openpyxl read-only values | ran/lost | `openpyxl-read-only-values` | 224.432 | 229.401 | 0.021 | 65.0 MiB | Streaming data-only value read. Lower memory; 25.44x slower by median. |
| pyopenxlsx | ran/lost | `pyopenxlsx` | 243.951 | 257.432 | 0.033 | 78.3 MiB | Cell materialization. 27.65x slower by median. |
| Polars calamine | not comparable | `polars-calamine` | 14.882 | 15.457 | 0.025 | 103.5 MiB | Not counted. Runner completed but reported semantic mismatch on this table-heavy workload. |
| fastexcel Python | not comparable | `fastexcel` | 15.211 | 16.135 | 0.050 | 66.4 MiB | Not counted. Runner completed but reported semantic mismatch on this table-heavy workload. |
| Polars xlsx2csv | not comparable | `polars-xlsx2csv` | 79.123 | 80.842 | 0.013 | 78.0 MiB | Not counted. Runner completed but reported semantic mismatch on this table-heavy workload. |
| Polars openpyxl | not comparable | `polars-openpyxl` | 206.852 | 223.053 | 0.075 | 120.7 MiB | Not counted. Runner completed but reported semantic mismatch on this table-heavy workload. |
| ClosedXML | blocked | `closedxml` | n/a | n/a | n/a | n/a | Not counted. The .NET runner failed with `CSSM_ModuleLoad(): One or more parameters passed to a function were not valid.` |
| fastxlsx | runner unavailable | `fastxlsx` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `fastxlsx` is not installed in the clean Python environment. |

In-process JavaScript reference:

| Library | Status | Median ms | P95 ms | CV | Peak RSS |
| --- | --- | ---: | ---: | ---: | ---: |
| Ascend in-process | ran/won | 8.594 | 8.784 | 0.056 | 209.5 MiB |
| SheetJS `xlsx@0.18.5` | ran/lost | 37.050 | 43.049 | 0.084 | 312.2 MiB |
| ExcelJS `4.4.0` | ran/lost | 53.997 | 61.667 | 0.069 | 324.0 MiB |

Coverage gate result: failed, as expected for a partial profile. The table-heavy row lacks ClosedXML and fastxlsx coverage, and several completed Python/Polars rows are semantic mismatches rather than losses. The profile remains missing `string-heavy` from current commit, `feature-rich`, `selected-sheet`, `metadata-only`, and `warm-workflow` coverage.

Humble allowed wording:

> On the generated `table-heavy` 2000x20 raw OOXML workload at commit `65520519`, Ascend's file-path value-read row had the fastest eligible median and p95 among completed comparable external runners. Several completed runners were semantic mismatches, and the full `xlsx-read-sota` profile remains incomplete.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend has a clean table-heavy read speed win across all runners."
- "Ascend beats fastexcel Python, Polars calamine, Polars xlsx2csv, or Polars openpyxl" from this run.
- "Ascend proves table metadata fidelity or preservation" from this value-read run.
- Any wording that treats semantic mismatches, blocked runners, unavailable runners, or the incomplete full profile as wins.

Next action: defer production optimization. Continue profile expansion with `feature-rich`; keep ClosedXML and fastxlsx as blockers unless their runners are fixed.

## Cycle: Feature-Rich Metadata Read

Classification: optimize, then validated ran/won. A narrow production optimization was justified and landed in `05656d4e`.

Workflow: XLSX open/inspect value read with rich workbook metadata for a feature-rich worksheet.

Why it matters for release: this is a required `xlsx-read-sota` workload and the release workflow that asks Ascend to inspect values while also proving it sees comments, hyperlinks, data validation, conditional formatting, and defined names.

Public/tracked-clean input: `competitive-io` generated `feature-rich` workbook using tracked benchmark code from detached commit `05656d4e`, `raw-ooxml` source, 2000 rows x 20 columns, 40,000 logical cells, 40,000 populated cells, 114,404 input bytes. No private corpus or local research workbook was used.

Optimization basis: the pre-optimization clean run at detached commit `222c4898` made the comparable Ascend rich-metadata row eligible but slower than Apache POI: `ascend-readxlsx-values-rich-metadata-bytes` median 92.937 ms, p95 147.127 ms, CV 0.313, peak RSS 171.1 MiB; Apache POI median 84.075 ms, p95 166.559 ms, CV 0.399, peak heap 344.0 MiB. Faster Ascend rows in that run were semantic mismatches for `read-values-rich-metadata` and were not counted.

Production change: `05656d4e` lets rich-metadata value reads use the existing byte sheet-data parser for scalar cells, then hydrate outer sheet metadata from stripped XML. Formula-containing sheets still fall back instead of being treated as byte-parser wins.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-05656d4e 05656d4e
cd /private/tmp/ascend-perf-hillclimb-05656d4e
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload feature-rich --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-05656d4e-runs/feature-rich-read-values.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-05656d4e-runs/feature-rich-read-values.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-05656d4e-runs/feature-rich-scoreboard.json
```

Environment:

- Commit: `05656d4e`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-05656d4e`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Python: `3.13.3`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-05656d4e-runs/feature-rich-read-values.json
/private/tmp/ascend-perf-hillclimb-05656d4e-runs/feature-rich-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. Rows marked `blocked`, `runner unavailable`, or `not comparable` are non-ranking status rows for claim wording and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend-readxlsx-values-rich-metadata-bytes` | 41.335 | 41.442 | 0.048 | 171.3 MiB | Comparable rich-metadata value read. Assertions saw 1 comment, 1 hyperlink, 1 data validation, 1 conditional format, 1 defined name, and `readFeatureRichMatches=true`. |
| Apache POI | ran/lost | `apache-poi` | 86.697 | 217.352 | 0.576 | 344.0 MiB heap | Materialized workbook rich-metadata read. Tail is noisy; bounded gap versus Ascend is 2.10x slower by median. |
| openpyxl | ran/lost | `openpyxl` | 129.402 | 146.646 | 0.107 | 103.3 MiB | Full openpyxl rich-metadata read. Lower memory; 3.13x slower by median. |
| Excelize | ran/lost | `excelize` | 134.029 | 172.187 | 0.130 | 46.0 MiB | File-path materialized rich-metadata read. Lower memory; 3.24x slower by median. |
| ExcelJS | ran/lost | `exceljs` | 51.325 | 64.856 | 0.114 | 315.1 MiB | In-process rich-metadata read. Compare only with Ascend in-process row, not external-process lanes. |
| SheetJS | not comparable | `sheetjs` | 33.529 | 36.453 | 0.066 | 297.1 MiB | Not counted. Runner completed but missed data validation and conditional formatting, so `readFeatureRichMatches=false`. |
| fastexcel Python | not comparable | `fastexcel` | 8.700 | 10.723 | 0.144 | 65.0 MiB | Not counted. Runner completed but did not report the required rich metadata. |
| python-calamine | not comparable | `python-calamine` | 15.892 | 16.310 | 0.022 | 42.0 MiB | Not counted. Runner completed but did not report the required rich metadata. |
| rust-calamine | not comparable | `rust-calamine` | 35.815 | 36.508 | 0.015 | 13.9 MiB | Not counted. Runner completed but did not report the required rich metadata. |
| FastExcel Java | not comparable | `fastexcel-java` | 11.019 | 28.226 | 0.572 | 112.0 MiB heap | Not counted. Runner completed but did not report the required rich metadata. |
| Polars calamine | not comparable | `polars-calamine` | 9.830 | 10.273 | 0.042 | 104.1 MiB | Not counted. Runner completed but did not report the required rich metadata. |
| Polars xlsx2csv | not comparable | `polars-xlsx2csv` | 74.243 | 74.744 | 0.005 | 76.2 MiB | Not counted. Runner completed but did not report the required rich metadata. |
| Polars openpyxl | not comparable | `polars-openpyxl` | 148.654 | 153.733 | 0.070 | 119.6 MiB | Not counted. Runner completed but did not report the required rich metadata. |
| pyopenxlsx | not comparable | `pyopenxlsx` | 230.839 | 234.438 | 0.016 | 74.3 MiB | Not counted. Runner completed but did not report the required rich metadata. |
| ClosedXML | blocked | `closedxml` | n/a | n/a | n/a | n/a | Not counted. The .NET runner failed with `CSSM_ModuleLoad(): One or more parameters passed to a function were not valid.` |
| fastxlsx | runner unavailable | `fastxlsx` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `fastxlsx` is not installed in the clean Python environment. |

In-process JavaScript reference:

| Library | Status | Median ms | P95 ms | CV | Peak RSS |
| --- | --- | ---: | ---: | ---: | ---: |
| Ascend in-process | ran/won | 5.472 | 9.091 | 0.263 | 198.2 MiB |
| ExcelJS `4.4.0` | ran/lost | 51.325 | 64.856 | 0.114 | 315.1 MiB |
| SheetJS `xlsx@0.18.5` | not comparable | 33.529 | 36.453 | 0.066 | 297.1 MiB |

Coverage gate result: failed, as expected for a partial profile. The feature-rich row still lacks ClosedXML coverage, and SheetJS/Calamine-family rows are semantic mismatches for rich metadata. The profile remains missing `string-heavy` from current commit, `selected-sheet`, `metadata-only`, and `warm-workflow` coverage.

Humble allowed wording:

> On the generated `feature-rich` 2000x20 raw OOXML workload at commit `05656d4e`, Ascend's rich-metadata value-read row was faster than the completed comparable external rich-metadata readers after the byte sheet-data parser optimization. Several faster runners were semantic mismatches and are not counted.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend beats SheetJS, fastexcel, Calamine, FastExcel Java, Polars, or pyopenxlsx" from this rich-metadata run.
- "Ascend beats ClosedXML or fastxlsx" from this run.
- Any wording that counts semantic mismatches, blocked runners, unavailable runners, or the incomplete full profile as wins.

Next action: optimize no further on this row. Continue profile expansion with `selected-sheet`; keep ClosedXML, fastxlsx, and rich-metadata semantic mismatches as external-baseline blockers unless their runners are fixed.

## Cycle: SheetJS Feature-Rich Rich Metadata

Classification: defer. The previous SheetJS feature-rich row was cooked as a
head-to-head because it used SheetJS' cell/comment/hyperlink surface but did not
enable the public `bookFiles` option needed to inspect the workbook package for
data validations and conditional formatting. This cycle fixes the runner
semantics and records SheetJS as a measured `ran/lost` row. No production
optimization is justified because Ascend wins the comparable in-process
feature-rich rich-metadata workflow.

Workflow: XLSX open/inspect value read with rich workbook metadata for a
feature-rich worksheet.

Why it matters for release: feature-rich inspection is the workflow behind
trustworthy agent decisions on comments, hyperlinks, data validations,
conditional formatting, and defined names. Counting a feature-dropping SheetJS
row as a win would be dishonest; making it comparable strengthens the claim.

Public/tracked-clean input: `competitive-io` generated the `feature-rich`
raw OOXML workload from detached commit `15119c8d`, 2000 rows x 20 columns,
40,000 logical cells, 40,000 populated cells, 114,404 input bytes. No private
corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-15119c8d 15119c8d
cd /private/tmp/ascend-perf-hillclimb-15119c8d
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --libraries ascend,sheetjs --workload feature-rich --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each > /private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-sheetjs-inprocess.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-sheetjs-inprocess.json --json --metric medianMs > /private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-sheetjs-inprocess-scoreboard.json
```

Environment:

- Commit: `15119c8d828e52493866c760c7fe28972e4a3bee`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-15119c8d`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Node: `24.3.0`
- SheetJS: `xlsx@0.18.5`
- Runtime profile: `category read`, `workload feature-rich`, `readSource raw-ooxml`, `validationMode each`, `repeat 5`, `warmup 1`.

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-sheetjs-inprocess.json
/private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-sheetjs-inprocess-scoreboard.json
```

Both rows below use 5 measured samples after 1 warmup and share
`in-process-generated-feature-rich`. Both rows assert 1 comment, 1 hyperlink,
1 data validation, 1 conditional format, 1 defined name, and
`readFeatureRichSemanticMatches: true`.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend` | 5.123 | 5.575 | 0.048 | 222.7 MiB | In-process rich-metadata value read with exact feature-rich semantic assertions. |
| SheetJS | ran/lost | `sheetjs` | 31.211 | 32.814 | 0.035 | 328.1 MiB | In-process SheetJS read with `bookFiles` package inspection for data validation and conditional formatting; 6.09x slower than Ascend by median. |
| Calamine-family runners | not comparable | n/a | n/a | n/a | n/a | n/a | Still do not expose the required rich metadata assertions in the current harness. |

Scoreboard result for the focused SheetJS run:

- `leaderFailures: []`
- `profileLeaderFailures: []`
- `coverageFailures: []`
- `coverageGaps: []`
- The `feature-rich` group winner was `ascend`.

Humble allowed wording:

> On the generated `feature-rich` raw OOXML workload at commit `15119c8d`, SheetJS `xlsx@0.18.5` was made semantically comparable by enabling `bookFiles` and checking the package metadata needed for comments, hyperlinks, data validations, conditional formatting, and defined names. Ascend was faster by median in this focused in-process comparison.

Forbidden wording:

- "Ascend beats Calamine-family readers on feature-rich rich-metadata reads."
- "SheetJS was always comparable in older feature-rich runs."
- Any wording that treats raw value-only Calamine-family rows as rich-metadata losses.

Next action: defer production optimization. Keep Calamine-family feature-rich
rich-metadata rows as not comparable unless their public APIs can expose the
required metadata, then assemble a current full-profile or merged
`xlsx-read-sota` gate using the cleaned selected-sheet, metadata-only,
FastXLSX, ClosedXML, and SheetJS evidence.

## Cycle: Calamine Feature-Rich Rich Metadata Boundary

Classification: not comparable/defer. No production optimization is justified.
The Calamine-family rows run, but their current public runner surfaces expose
values, sheet names, dimensions, and merged ranges, not the feature-rich metadata
required for this release workflow.

Workflow: XLSX open/inspect value read with rich workbook metadata for a
feature-rich worksheet.

Why it matters for release: Calamine-family readers are fast value readers, but
the release claim must not count value-only speed as a win on a workflow that
requires comments, hyperlinks, data validations, conditional formatting, and
defined names.

Public/tracked-clean input: `competitive-io` generated the `feature-rich`
raw OOXML workload from detached commit `15119c8d`, 2000 rows x 20 columns,
40,000 logical cells, 40,000 populated cells, 114,404 input bytes. No private
corpus or local research workbook was used.

Commands:

```bash
cd /private/tmp/ascend-perf-hillclimb-15119c8d
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-readxlsx-values-rich-metadata-bytes,python-calamine,rust-calamine --workload feature-rich --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-calamine-boundary.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-calamine-boundary.json --json --metric medianMs > /private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-calamine-boundary-scoreboard.json
```

Environment:

- Commit: `15119c8d828e52493866c760c7fe28972e4a3bee`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-15119c8d`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Node: `24.3.0`
- Runtime profile: `category read`, `executionScope external-process`, `workload feature-rich`, `readSource raw-ooxml`, `validationMode each`, `repeat 5`, `warmup 1`.

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-calamine-boundary.json
/private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-calamine-boundary-scoreboard.json
```

All rows below use 5 measured samples after 1 warmup.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won only in its own lane | `ascend-readxlsx-values-rich-metadata-bytes` | 56.250 | 57.118 | 0.019 | 161.2 MiB | Rich-metadata assertions pass, but this preloaded-bytes lane is not the same timing lane as Calamine materialization. |
| python-calamine | not comparable | `python-calamine` | 24.332 | 39.131 | 0.246 | 42.4 MiB | Runner completed but reported 0 comments, 0 hyperlinks, 0 data validations, 0 conditional formats, and 0 defined names; status `semantic-mismatch`. |
| rust-calamine | not comparable | `rust-calamine` | 135.078 | 162.012 | 0.280 | 15.6 MiB | Runner completed but reported 0 comments, 0 hyperlinks, 0 data validations, 0 conditional formats, and 0 defined names; status `semantic-mismatch`. |

Scoreboard result for the focused Calamine boundary run:

- `leaderFailures: []`
- `profileLeaderFailures: []`
- `coverageFailures: []`
- `coverageGaps: []`
- The Calamine timing lane has no winner because both rows are
  `semantic-mismatch`.

Humble allowed wording:

> Python Calamine and Rust Calamine completed the generated `feature-rich` read workflow, but their current benchmark runners did not expose the rich metadata required for the release claim. They are not comparable for rich-metadata feature-rich reads and are not counted as Ascend wins.

Forbidden wording:

- "Ascend beats Calamine on feature-rich rich-metadata reads."
- "Calamine lost the feature-rich rich-metadata benchmark."
- Any wording that counts value-only Calamine timings as rich-metadata evidence.

Next action: kill Calamine-family rich-metadata speed comparisons unless a
public Calamine API can expose the required metadata. Continue by assembling the
current full-profile or merged `xlsx-read-sota` gate from accepted comparable
evidence and explicit not-comparable boundaries.

## Cycle: Selected-Sheet Value Read

Classification: defer. No production optimization is justified from this cycle.

Workflow: XLSX open/inspect value read for one selected worksheet out of a multi-sheet workbook.

Why it matters for release: this is a required `xlsx-read-sota` workload and a common agent workflow when a user asks to inspect or operate on one named sheet without hydrating the rest of the workbook.

Public/tracked-clean input: `competitive-io` generated `selected-sheet` workbook using tracked benchmark code from detached commit `5055d794`, `raw-ooxml` source, 2000 rows x 20 columns in the selected `Data` sheet, 40,000 logical cells, 40,000 populated cells, 114,747 input bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-5055d794 5055d794
cd /private/tmp/ascend-perf-hillclimb-5055d794
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload selected-sheet --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-5055d794-runs/selected-sheet-read-values.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-5055d794-runs/selected-sheet-read-values.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-5055d794-runs/selected-sheet-scoreboard.json
```

Environment:

- Commit: `5055d794`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-5055d794`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Python: `3.13.3`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-5055d794-runs/selected-sheet-read-values.json
/private/tmp/ascend-perf-hillclimb-5055d794-runs/selected-sheet-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. Unsupported-operation rows are non-ranking status rows for claim wording and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend` | 3.054 | 3.156 | 0.042 | 198.0 MiB | Selected-sheet value read. Assertions loaded only `Data`, source sheet count 3, loaded sheet count 1, and `selectedSheetMatches=true`. |
| SheetJS | ran/lost | `sheetjs` | 29.149 | 30.395 | 0.039 | 290.7 MiB | Selected-sheet value read. Bounded gap versus Ascend is 9.54x slower by median. |
| ExcelJS | not comparable | n/a | n/a | n/a | n/a | n/a | Unsupported operation for this profile; not counted. |
| openpyxl | not comparable | n/a | n/a | n/a | n/a | n/a | Unsupported operation for this profile; not counted. |
| Calamine | not comparable | n/a | n/a | n/a | n/a | n/a | Unsupported operation for this profile; not counted. |
| Apache POI | not comparable | n/a | n/a | n/a | n/a | n/a | Unsupported operation for this profile; not counted. |
| ClosedXML | not comparable | n/a | n/a | n/a | n/a | n/a | Unsupported operation for this profile; not counted. |

Coverage gate result: failed, as expected for a partial profile. The selected-sheet row is only comparable against SheetJS in the recorded clean profile; ExcelJS, openpyxl, Calamine, Apache POI, and ClosedXML were explicit unsupported-operation gaps in that historical run. The profile remains missing `string-heavy` from current commit, `metadata-only`, and `warm-workflow` coverage.

Supersession note: the OpenPyXL unsupported-operation status above is historical
for the clean detached `5055d794` cycle. Current harness tests now prove an
OpenPyXL selected-sheet projection row can run and pass semantic assertions.
Do not route current owner work as "make OpenPyXL selected-sheet run"; route it
as "record a clean repeat-5 selected-sheet benchmark and resolve the timing-lane
boundary before promoting any OpenPyXL selected-sheet speed wording."

Supersession note: the Calamine unsupported-operation status above is also
historical for the clean detached `5055d794` cycle. Commit `79d6cefd` proves the
python-calamine runner can project only the `Data` sheet and join the same
selected-sheet external timing lane. Do not route current owner work as "make
Calamine selected-sheet run"; route it as "record a clean repeat-5 selected-sheet
benchmark and resolve the broad-claim blockers before promoting any Calamine
selected-sheet speed wording."

Humble allowed wording:

> On the generated `selected-sheet` raw OOXML workload at commit `5055d794`, Ascend loaded the selected `Data` sheet faster than the only completed comparable external selected-sheet runner, SheetJS. Other libraries in the profile are unsupported for this operation and are not counted as wins.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend beats ExcelJS, openpyxl, Calamine, Apache POI, or ClosedXML" from this selected-sheet run.
- Any wording that counts unsupported operations or the incomplete full profile as wins.

Next action: defer production optimization. Continue profile expansion with `metadata-only`.

## Fold-In: OpenPyXL Selected-Sheet Projection

Classification: accepted evidence plus benchmark blocker. This closes the
current OpenPyXL generated selected-sheet `unsupported-operation` gap in the
harness, but it does not promote a speed claim.

Workflow: generated `selected-sheet` read, projecting only the `Data` worksheet
from a workbook that also contains `Summary` and `Archive`.

Why it matters for release: selected-sheet open/inspect is a practical
agent-native workflow for unknown workbooks. The release proof must distinguish
"can produce the selected-sheet semantic view" from "proved fastest under the
same timing boundary" and from "proved true internal one-sheet hydration."

Evidence we have:

- `fixtures/benchmarks/runners/openpyxl_runner.py` accepts `--selected-sheet`
  for read operations and reports `selectedSheetRead: true`,
  `sourceSheetCount: 3`, `loadedSheetCount: 1`, `loadedSheetNames: Data`,
  `hasAllSheets: false`, and selected-sheet cell/formula/range hashes.
- `fixtures/benchmarks/competitive-io.ts` includes external selected-sheet
  read cases only for runners declaring `selectedSheetRead: true`, and passes
  `--selected-sheet Data` to those runners.
- `fixtures/benchmarks/runners/openpyxl.manifest.json`,
  `fixtures/benchmarks/runners/python-readers.manifest.json`, and
  `fixtures/benchmarks/runners/ascend-python-readers.manifest.json` declare
  `selectedSheetRead: true` for the non-read-only OpenPyXL runner.
- Focused validation passed:
  `python3 -m py_compile fixtures/benchmarks/runners/openpyxl_runner.py`,
  `bun test fixtures/benchmarks/competitive-io.test.ts -t "selected-sheet" --timeout 30000`,
  and `bun test fixtures/benchmarks/competitive-real-workbook.test.ts -t "external runner manifests preserve normalized metadata|combined reader manifests include direct rust calamine and excelize coverage"`.
- A current-worktree bounded run at `/private/tmp/ascend-selected-sheet-openpyxl-current.json`
  recorded pass rows for Ascend, SheetJS, and OpenPyXL. OpenPyXL reported
  median `204.707 ms`, p95 `222.561 ms`, CV `0.070`, and peak RSS `86.4 MiB`
  for its external selected-sheet projection row; Ascend and SheetJS remained
  in the in-process selected-sheet lane at medians `3.172 ms` and `27.188 ms`.

Evidence missing:

- No clean detached repeat-5 post-fold-in selected-sheet profile has been
  recorded in this matrix.
- The current scoreboard separates OpenPyXL's
  `external-internal-operation-timing:selected-sheet` lane from Ascend and
  SheetJS `in-process-generated-selected-sheet`, so this evidence does not
  create a fair Ascend-vs-OpenPyXL speed comparison.
- OpenPyXL is proven here as selected-sheet projection semantics, not as true
  preservation-first one-sheet package hydration. Do not claim it avoids loading
  other workbook internals.
- ExcelJS, Apache POI, and ClosedXML selected-sheet gaps remain; Calamine
  selected-sheet now has current-worktree runner proof but no accepted clean
  repeat-5 benchmark row in this matrix.

Competitor/QSS contrast: this reduces an OpenPyXL unsupported-operation gap in
the external baseline matrix, but it does not change the QSS/SOTA decision.
Ascend still cannot make a broad speed claim, and OpenPyXL timing must stay in
its own boundary until the benchmark loop records a comparable external Ascend
selected-sheet row or explicitly downgrades the comparison.

Allowed wording:

> Ascend's benchmark harness can now run a generated selected-sheet semantic
> projection against OpenPyXL and verify that only the requested `Data` sheet is
> represented in the output assertions. A clean timing-boundary rerun is still
> required before any OpenPyXL selected-sheet speed wording is allowed.

Forbidden wording:

- "OpenPyXL selected-sheet is still unsupported" as a current harness blocker.
- "Ascend beats OpenPyXL for selected-sheet reads" from the current worktree
  run.
- "OpenPyXL hydrates only one sheet internally."
- Any broad XLSX read, SOTA, or QSS-leapfrog wording from this fold-in.

Next owner action: benchmarking/external baselines owns the clean rerun and
timing-boundary decision.

```bash
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --libraries ascend,sheetjs,openpyxl --workload selected-sheet --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/openpyxl.manifest.json > /private/tmp/ascend-selected-sheet-openpyxl-clean.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-selected-sheet-openpyxl-clean.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-selected-sheet-openpyxl-clean-scoreboard.json
```

Acceptance evidence: clean worktree status, exact commit, environment versions,
OpenPyXL runner/library version, all three selected-sheet semantic assertions
passing, and an explicit scoreboard decision that either adds a comparable
external Ascend selected-sheet row or keeps Ascend-vs-OpenPyXL selected-sheet
speed wording forbidden because the timing lanes differ.

## Cycle: Metadata-Only Read

Classification: defer. No production optimization is justified from this cycle.

Workflow: XLSX metadata-only open/inspect without hydrating cells.

Why it matters for release: this is a required `xlsx-read-sota` workload and the fastest safe first-pass workflow when an agent needs sheet inventory before deciding whether to hydrate workbook values.

Public/tracked-clean input: `competitive-io` generated `metadata-only` workbook using tracked benchmark code from detached commit `5261b08d`, `raw-ooxml` source, 200 rows x 20 columns, 4,000 logical cells, 4,000 populated cells, 15,347 input bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-5261b08d 5261b08d
cd /private/tmp/ascend-perf-hillclimb-5261b08d
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload metadata-only --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-5261b08d-runs/metadata-only-read-values.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-5261b08d-runs/metadata-only-read-values.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-5261b08d-runs/metadata-only-scoreboard.json
```

Environment:

- Commit: `5261b08d`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-5261b08d`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Python: `3.13.3`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-5261b08d-runs/metadata-only-read-values.json
/private/tmp/ascend-perf-hillclimb-5261b08d-runs/metadata-only-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. Unsupported-operation rows are non-ranking status rows for claim wording and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend external | ran/won | `ascend-external-metadata-only-bytes` | 0.310 | 0.329 | 0.082 | 89.8 MiB | External metadata-only read over preloaded bytes. Assertions loaded 3 sheets, hydrated no cells, and `cellsNotHydrated=true`. |
| openpyxl | ran/lost | `openpyxl-metadata-only` | 1.986 | 3.492 | 0.296 | 50.8 MiB | Metadata-only workbook inventory. Lower memory; tail is moderately noisy; 6.41x slower by median. |
| Ascend in-process | ran/won | `ascend` | 0.325 | 0.375 | 0.109 | 152.1 MiB | In-process metadata-only read. Compare with SheetJS in-process, not external-process lanes. |
| SheetJS | ran/lost | `sheetjs` | 0.832 | 0.892 | 0.056 | 154.5 MiB | In-process metadata-only read. 2.56x slower than Ascend in-process by median. |
| ExcelJS | not comparable | n/a | n/a | n/a | n/a | n/a | Unsupported operation for this profile; not counted. |
| Calamine | not comparable | n/a | n/a | n/a | n/a | n/a | Unsupported operation for this profile; not counted. |
| Apache POI | not comparable | n/a | n/a | n/a | n/a | n/a | Unsupported operation for this profile; not counted. |
| ClosedXML | not comparable | n/a | n/a | n/a | n/a | n/a | Unsupported operation for this profile; not counted. |

Coverage gate result: failed, as expected for a partial profile. The metadata-only row has comparable Ascend, SheetJS, and openpyxl evidence, while ExcelJS, Calamine, Apache POI, and ClosedXML are explicit unsupported-operation gaps. The profile remains missing `string-heavy` from current commit and `warm-workflow` coverage.

Humble allowed wording:

> On the generated `metadata-only` raw OOXML workload at commit `5261b08d`, Ascend's metadata-only readers were faster than the completed comparable SheetJS and openpyxl rows while proving cells were not hydrated. Unsupported metadata-only operations in other libraries are not counted as wins.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend beats ExcelJS, Calamine, Apache POI, or ClosedXML" from this metadata-only run.
- Any wording that counts unsupported operations or the incomplete full profile as wins.

Next action: defer production optimization. Continue profile expansion with `warm-workflow`.

## Cycle: Metadata-Only Same-Lane External Read

Classification: defer. The previous metadata-only row mixed a preloaded-bytes
Ascend external lane with in-process SheetJS and external openpyxl timing lanes.
This cycle fixes the evidence boundary by running Ascend, SheetJS, and openpyxl
through one external-process metadata-only load lane. No production optimization
is justified because Ascend wins the comparable rows.

Workflow: XLSX metadata-only open/inspect without hydrating cells.

Why it matters for release: this is the fastest safe first-pass workflow for an
agent that needs workbook sheet inventory before deciding which cells to hydrate.
It is also a required `xlsx-read-sota` coverage row.

Public/tracked-clean input: `competitive-io` generated the `metadata-only` raw
OOXML workload from detached commit `fa3a13dc`, 200 rows x 20 columns, 4,000
logical cells, 4,000 populated cells, 15,347 input bytes. No private corpus or
local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-fa3a13dc fa3a13dc
cd /private/tmp/ascend-perf-hillclimb-fa3a13dc
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only,sheetjs-metadata-only,openpyxl-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-fa3a13dc-runs/metadata-only-same-lane.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-fa3a13dc-runs/metadata-only-same-lane.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-fa3a13dc-runs/metadata-only-same-lane-scoreboard.json
```

Environment:

- Commit: `fa3a13dc1f72de489d6d301bf1f81cbe3400df0f`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-fa3a13dc`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Node: `22.22.0`
- Python: `3.9.6`
- SheetJS: `xlsx@0.18.5`
- openpyxl: `3.1.5`
- Runtime profile: `category read`, `executionScope external-process`, `workload metadata-only`, `readSource raw-ooxml`, `validationMode each`, `repeat 5`, `warmup 1`.

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-fa3a13dc-runs/metadata-only-same-lane.json
/private/tmp/ascend-perf-hillclimb-fa3a13dc-runs/metadata-only-same-lane-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. All
three rows share `external-internal-metadata-only-load-timing:metadata-only` and
assert `metadataOnlyRead: true`, `sourceSheetCount: 3`, `loadedSheetCount: 3`,
`loadedSheetNames: Data,Summary,Archive`, and `cellsHydrated: false`.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend | ran/won | `ascend-external-metadata-only` | 0.394 | 0.457 | 0.148 | 91.2 MiB | External-process metadata-only read from file path; same timing lane as SheetJS and openpyxl; no cells hydrated. |
| SheetJS | ran/lost | `sheetjs-metadata-only` | 0.941 | 1.106 | 0.110 | 148.8 MiB | External-process SheetJS `bookSheets` metadata load; 2.39x slower than Ascend by median. |
| openpyxl | ran/lost | `openpyxl-metadata-only` | 2.050 | 3.284 | 0.257 | 50.9 MiB | External-process read-only metadata inventory; lower memory than Ascend but 5.21x slower by median and noisier. |
| ExcelJS | not comparable | n/a | n/a | n/a | n/a | n/a | Still skipped: metadata-only read is unsupported by this harness. |
| Calamine | not comparable | n/a | n/a | n/a | n/a | n/a | Still a profile `unsupported-operation` gap. |
| Apache POI | not comparable | n/a | n/a | n/a | n/a | n/a | Still a profile `unsupported-operation` gap. |
| ClosedXML | not comparable | n/a | n/a | n/a | n/a | n/a | Still a profile `unsupported-operation` gap. |

Scoreboard result for the focused same-lane run:

- `leaderFailures: []`
- `profileLeaderFailures: []`
- No metadata-only `coverageFailures` remain for Ascend, SheetJS, or openpyxl.
- Metadata-only `coverageGaps` remain for ExcelJS, Calamine, Apache POI, and
  ClosedXML because their metadata-only operation is unsupported in the current
  profile.

Humble allowed wording:

> On the generated `metadata-only` raw OOXML workload at commit `fa3a13dc`, Ascend's external-process metadata-only read was faster by median than same-lane SheetJS and openpyxl rows while proving that cells were not hydrated. ExcelJS, Calamine, Apache POI, and ClosedXML remain unsupported-operation gaps, so this is scoped metadata-only evidence, not a broad XLSX-read claim.

Forbidden wording:

- "Ascend has a full metadata-only SOTA claim across every library."
- "Ascend beats ExcelJS, Calamine, Apache POI, or ClosedXML" from this metadata-only run.
- Any wording that treats unsupported metadata-only competitors as wins.

Next action: defer production optimization. The next highest-impact blockers are
feature-rich semantic mismatches for SheetJS and Calamine, the unavailable
fastxlsx runner, and assembling a current full-profile gate from the cleaned
same-lane evidence.

## Cycle: FastXLSX Cell-Materialization Coverage

Classification: defer. The historical FastXLSX row was a cooked non-result
because the clean Python environment lacked `fastxlsx`. This cycle installs
`fastxlsx==0.2.0` into an isolated Python 3.12 target under `/private/tmp`,
runs it through the existing tracked runner, and compares it only against
Ascend's same `external-internal-cell-materialization-timing` lane. No
production optimization is justified because Ascend wins every comparable
value/warm row in this lane.

Workflow: XLSX open/inspect value read with cell materialization.

Why it matters for release: FastXLSX is a public high-performance Python XLSX
reader/writer candidate. Converting it from `runner unavailable` to measured
same-lane evidence prevents a weak Ascend claim from hiding a missing competitor.

Public/tracked-clean input: `competitive-io` generated `workload all` raw OOXML
inputs from detached commit `52f7f172`. The required release rows were
`dense-values`, `sparse-wide`, `string-heavy`, `styles-heavy`, `formula-heavy`,
`table-heavy`, `feature-rich`, and `warm-workflow`; the harness also emitted
non-profile generated read workloads. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-52f7f172 52f7f172
cd /private/tmp/ascend-perf-hillclimb-52f7f172
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
/opt/homebrew/bin/python3.12 -m pip install --target /private/tmp/ascend-fastxlsx-py312 fastxlsx==0.2.0 openpyxl psutil
mkdir -p /private/tmp/ascend-py312-bin
ln -sf /opt/homebrew/bin/python3.12 /private/tmp/ascend-py312-bin/python3
PYTHONPATH=/private/tmp/ascend-fastxlsx-py312 TMPDIR=/private/tmp env PATH=/private/tmp/ascend-py312-bin:/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-readxlsx-cell-materialization-bytes,fastxlsx --workload all --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-52f7f172-runs/fastxlsx-cell-materialization-all.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-52f7f172-runs/fastxlsx-cell-materialization-all.json --json --metric medianMs > /private/tmp/ascend-perf-hillclimb-52f7f172-runs/fastxlsx-cell-materialization-all-scoreboard.json
```

Environment:

- Commit: `52f7f172e2826211a4a0ba811a722077dcd1a824`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-52f7f172`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Node: `24.3.0`
- Python runner: `/opt/homebrew/bin/python3.12` exposed as `/private/tmp/ascend-py312-bin/python3`
- FastXLSX: `0.2.0`
- Runtime profile: `category read`, `executionScope external-process`, `workload all`, `readSource raw-ooxml`, `validationMode each`, `repeat 5`, `warmup 1`.

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-52f7f172-runs/fastxlsx-cell-materialization-all.json
/private/tmp/ascend-perf-hillclimb-52f7f172-runs/fastxlsx-cell-materialization-all-scoreboard.json
```

All comparable rows below use 5 measured samples after 1 warmup and share
`external-internal-cell-materialization-timing:<workload>`. FastXLSX rows report
`runnerVersion: 0.2.0`. Peak RSS is reported in MiB.

| Workload | Ascend status | Ascend median / p95 / CV / RSS | FastXLSX status | FastXLSX median / p95 / CV / RSS | Semantic comparability |
| --- | --- | ---: | --- | ---: | --- |
| `dense-values` | ran/won | 9.270 ms / 10.247 ms / 0.067 / 102.6 | ran/lost | 34.884 ms / 57.856 ms / 0.287 / 52.5 | Same cell-materialization value semantics; FastXLSX uses less RSS but is 3.76x slower by median. |
| `sparse-wide` | ran/won | 13.066 ms / 14.376 ms / 0.047 / 127.3 | ran/lost | 101.486 ms / 102.578 ms / 0.009 / 87.5 | Same value semantics; FastXLSX is 7.77x slower by median. |
| `string-heavy` | ran/won | 14.391 ms / 14.661 ms / 0.027 / 125.7 | ran/lost | 27.171 ms / 27.750 ms / 0.013 / 57.4 | Same value semantics; FastXLSX is 1.89x slower by median. |
| `styles-heavy` | ran/won | 12.420 ms / 24.971 ms / 0.379 / 110.2 | ran/lost | 27.336 ms / 28.354 ms / 0.019 / 54.3 | Same value semantics; Ascend tail is noisy but still faster by median and p95. |
| `formula-heavy` | ran/won | 11.562 ms / 14.741 ms / 0.137 / 111.0 | ran/lost | 29.723 ms / 30.831 ms / 0.023 / 54.7 | Same cached-value semantics; FastXLSX is 2.57x slower by median. |
| `table-heavy` | ran/won | 14.155 ms / 14.259 ms / 0.028 / 124.5 | ran/lost | 26.208 ms / 26.715 ms / 0.010 / 55.3 | Same value semantics; table metadata is not part of this value lane. |
| `warm-workflow` | ran/won | 8.053 ms / 8.411 ms / 0.024 / 104.2 | ran/lost | 25.877 ms / 26.492 ms / 0.025 / 52.9 | Same value semantics on the warm workflow row; FastXLSX is 3.21x slower by median. |
| `feature-rich` | not comparable | 11.573 ms / 11.793 ms / 0.017 / 116.6 | not comparable | 25.734 ms / 25.911 ms / 0.006 / 53.3 | Both rows are `semantic-mismatch` for `read-values-rich-metadata`; not counted as wins or losses. |

Scoreboard result for the focused FastXLSX lane:

- `leaderFailures: []`
- `profileLeaderFailures: []`
- `coverageFailures: []`
- `coverageGaps: []`
- Every comparable FastXLSX value/warm group winner was
  `ascend-readxlsx-cell-materialization-bytes`.
- The `feature-rich` group has no winner because both rows are
  `semantic-mismatch` for rich metadata.

Humble allowed wording:

> In an isolated Python 3.12 environment with `fastxlsx==0.2.0`, FastXLSX successfully ran the generated cell-materialization value/warm workloads. Ascend's same-lane cell-materialization row was faster by median on every comparable value/warm row, while FastXLSX generally used less RSS. The feature-rich rich-metadata row remains not comparable and is not counted as a win.

Forbidden wording:

- "Ascend beats FastXLSX on rich-metadata feature-rich reads."
- "Ascend beats FastXLSX in every XLSX workflow."
- Any wording that hides FastXLSX's lower memory footprint or the isolated Python 3.12 dependency setup.

Next action: defer production optimization. Carry the isolated Python 3.12
FastXLSX setup into the next full-profile gate, and keep attacking
feature-rich rich-metadata semantic mismatches rather than optimizing this
winning value-read lane.

## Cycle: Current FastXLSX Carry-Forward Gate

Classification: comparable external evidence plus defer. FastXLSX is runnable
again in the isolated Python 3.12 environment at current commit `248f76d9`. The
focused all-workload run reproduced same-lane FastXLSX coverage and exposed a
noisy table-heavy median blip; a repeat-15 table-heavy rerun resolved it in
Ascend's favor. No production optimization is justified.

Workflow: XLSX open/inspect value read with cell materialization.

Why it matters for release: this carries the previously isolated FastXLSX setup
forward to the current claim state. It prevents stale `runner unavailable`
wording from being used where FastXLSX can now run, while still keeping the
comparison scoped to value materialization rather than rich metadata.

Public/tracked-clean input: `competitive-io` generated `workload all`
`raw-ooxml` inputs from tracked benchmark code at commit `248f76d9`. The
follow-up table-heavy run used the same generated workload shape, 2000 rows x 20
columns. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-fastxlsx-current-248f76d9 248f76d9
cd /private/tmp/ascend-fastxlsx-current-248f76d9
bun install --frozen-lockfile
/opt/homebrew/bin/python3.12 -m pip install --target /private/tmp/ascend-fastxlsx-py312 fastxlsx==0.2.0 openpyxl psutil
mkdir -p /private/tmp/ascend-py312-bin
ln -sf /opt/homebrew/bin/python3.12 /private/tmp/ascend-py312-bin/python3
PYTHONPATH=/private/tmp/ascend-fastxlsx-py312 TMPDIR=/private/tmp env PATH=/private/tmp/ascend-py312-bin:/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-readxlsx-cell-materialization-bytes,fastxlsx --workload all --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-cell-materialization-all.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-cell-materialization-all.json --json --metric medianMs --assert-leader ascend > /private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-cell-materialization-all-assert-ascend.json
PYTHONPATH=/private/tmp/ascend-fastxlsx-py312 TMPDIR=/private/tmp env PATH=/private/tmp/ascend-py312-bin:/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-readxlsx-cell-materialization-bytes,fastxlsx --workload table-heavy --read-source raw-ooxml --repeat 15 --warmup 3 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-table-heavy-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-table-heavy-repeat15.json --json --metric medianMs --assert-leader ascend > /private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-table-heavy-repeat15-scoreboard.json
```

Environment:

- Commit: `248f76d9f9812f2ac5106cf7b492e448cc8a11de`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-fastxlsx-current-248f76d9`
- Bun: `1.3.13`
- Python runner: `/opt/homebrew/bin/python3.12` exposed as
  `/private/tmp/ascend-py312-bin/python3`
- Python: `3.12.13`
- FastXLSX: `0.2.0`
- Runtime profile: `category read`, `executionScope external-process`,
  `readSource raw-ooxml`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-cell-materialization-all.json
/private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-cell-materialization-all-assert-ascend.json
/private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-table-heavy-repeat15.json
/private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-table-heavy-repeat15-scoreboard.json
```

Focused all-workload result, repeat 5 after 1 warmup, same
`external-internal-cell-materialization-timing:<workload>` lane:

| Workload | Ascend status | Ascend median / p95 / CV / RSS | FastXLSX status | FastXLSX median / p95 / CV / RSS | Semantic comparability |
| --- | --- | ---: | --- | ---: | --- |
| `dense-values` | ran/won | 19.885 ms / 25.612 ms / 0.164 / 101.3 MiB | ran/lost | 42.789 ms / 51.235 ms / 0.092 / 56.6 MiB | Same cell-materialization value semantics; FastXLSX uses less RSS but is 2.15x slower by median. |
| `string-heavy` | ran/won | 41.127 ms / 50.171 ms / 0.180 / 124.1 MiB | ran/lost | 103.684 ms / 119.541 ms / 0.168 / 58.5 MiB | Same value semantics; FastXLSX is 2.52x slower by median. |
| `sparse-wide` | ran/won | 41.422 ms / 63.959 ms / 0.273 / 127.3 MiB | ran/lost | 200.662 ms / 214.111 ms / 0.051 / 87.8 MiB | Same value semantics; FastXLSX is 4.84x slower by median. |
| `styles-heavy` | ran/won | 21.742 ms / 27.272 ms / 0.151 / 113.4 MiB | ran/lost | 49.134 ms / 51.950 ms / 0.042 / 55.0 MiB | Same value semantics; FastXLSX is 2.26x slower by median. |
| `formula-heavy` | ran/won | 25.768 ms / 72.697 ms / 0.604 / 109.7 MiB | ran/lost | 60.932 ms / 65.558 ms / 0.042 / 54.4 MiB | Same cached-value semantics; Ascend tail is noisy but median remains 2.36x faster. |
| `warm-workflow` | ran/won | 25.409 ms / 41.971 ms / 0.302 / 101.3 MiB | ran/lost | 51.711 ms / 54.289 ms / 0.078 / 53.4 MiB | Same value semantics; FastXLSX is 2.04x slower by median. |
| `feature-rich` | not comparable | 25.415 ms / 29.076 ms / 0.093 / 114.9 MiB | not comparable | 58.247 ms / 67.872 ms / 0.084 / 53.2 MiB | Both rows are `semantic-mismatch` for `read-values-rich-metadata`; not counted as wins or losses. |

Table-heavy repeat-15 check:

| Runner | Status | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `ascend-readxlsx-cell-materialization-bytes` | ran/won | 30.133 | 38.316 | 0.152 | 140.5 MiB | Same table-heavy value materialization; table metadata is not part of this lane. |
| `fastxlsx` | ran/lost | 46.843 | 59.085 | 0.106 | 55.6 MiB | Same value materialization; lower memory but 1.55x slower by median. |

Scoreboard result:

- Focused all-workload `--assert-leader ascend`: `leaderFailures: []`.
- Table-heavy repeat-15 `--assert-leader ascend`: `leaderFailures: []`.
- The all-workload table-heavy repeat-5 group briefly had `fastxlsx` as the
  median winner by 1.7%, but the difference was not a significant leader loss
  and the repeat-15 check reversed it.
- FastXLSX consistently used lower RSS than Ascend on this lane.
- The feature-rich rich-metadata row remains not comparable.

Humble allowed wording:

> In an isolated Python 3.12 environment with `fastxlsx==0.2.0` at current commit
> `248f76d9`, FastXLSX ran the generated same-lane cell-materialization
> value/warm workloads. Ascend had no significant leader failures and the
> repeat-15 table-heavy check favored Ascend, while FastXLSX used less memory.
> Feature-rich rich-metadata remains not comparable.

Forbidden wording:

- "Ascend beats FastXLSX on feature-rich rich-metadata reads."
- "Ascend beats FastXLSX on memory."
- "Ascend beats FastXLSX in every XLSX workflow."

Next action: defer production optimization on FastXLSX value materialization.
Carry this current FastXLSX setup into the next full-profile gate and keep
attacking feature-rich rich-metadata semantic mismatches or the metadata-only
Calamine loss.

## Cycle: Warm Workflow Value Read

Classification: defer. No production optimization is justified from this cycle.

Workflow: XLSX warm open/inspect value read.

Why it matters for release: this is a required `xlsx-read-sota` workload and represents repeated agent inspection after the runtime and parsers are already warm.

Public/tracked-clean input: `competitive-io` generated `warm-workflow` workbook using tracked benchmark code from detached commit `add13c79`, `raw-ooxml` source, 2000 rows x 20 columns, 40,000 logical cells, 40,000 populated cells, 112,699 input bytes. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-perf-hillclimb-add13c79 add13c79
cd /private/tmp/ascend-perf-hillclimb-add13c79
bun install --frozen-lockfile
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --workload warm-workflow --read-source raw-ooxml --repeat 5 --warmup 1 --validation-mode each --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json > /private/tmp/ascend-perf-hillclimb-add13c79-runs/warm-workflow-read-values.json
env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-perf-hillclimb-add13c79-runs/warm-workflow-read-values.json --json --metric medianMs --require-profile xlsx-read-sota > /private/tmp/ascend-perf-hillclimb-add13c79-runs/warm-workflow-scoreboard.json
```

Environment:

- Commit: `add13c79`
- Worktree: clean detached worktree at `/private/tmp/ascend-perf-hillclimb-add13c79`; `git status --short --branch` reported `## HEAD (no branch)` with no changed paths after the run.
- OS: Darwin 25.4.0 arm64
- Bun: `1.3.13`
- Python: `3.13.3`
- Cargo: `1.91.1`
- Maven: `3.9.15`, Java runtime `25.0.2` as reported by Maven
- .NET: `8.0.125`
- Go: `go1.26.3 darwin/arm64`

Raw output:

```text
/private/tmp/ascend-perf-hillclimb-add13c79-runs/warm-workflow-read-values.json
/private/tmp/ascend-perf-hillclimb-add13c79-runs/warm-workflow-scoreboard.json
```

All successful timing rows below use 5 measured samples after 1 warmup. Rows marked `blocked` or `runner unavailable` are non-ranking status rows for claim wording and are not counted as wins.

| Competitor | Status | Representative row | Median ms | P95 ms | CV | Peak RSS/heap | Semantic comparability |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| Ascend operation bytes | ran/won | `ascend-readxlsx-raw-values-operation-bytes` | 3.244 | 3.297 | 0.016 | 92.5 MiB | Warm value-read operation over preloaded bytes with ordered semantic assertions. Comparable for operation-timing warm workflow only. |
| Ascend operation path | ran/won | `ascend-readxlsx-raw-values-operation-path` | 3.773 | 3.905 | 0.042 | 91.3 MiB | Warm value-read operation from file path. Comparable for operation-timing warm workflow only. |
| FastExcel Java | ran/lost | `fastexcel-java` | 14.338 | 38.646 | 0.514 | 344.0 MiB heap | Streaming value read. Tail is noisy; 4.42x slower than Ascend operation-bytes by median. |
| python-calamine | ran/lost | `python-calamine` | 16.621 | 50.326 | 0.657 | 42.9 MiB | File-path materialized value read. Lower memory; tail is noisy; 5.12x slower by median. |
| Polars calamine | ran/lost | `polars-calamine` | 17.451 | 38.641 | 0.442 | 104.6 MiB | Value read through Polars/calamine. Tail is noisy; 5.38x slower by median. |
| rust-calamine | ran/lost | `rust-calamine` | 42.163 | 51.009 | 0.108 | 12.2 MiB | File-path materialized value read. Much lower memory; 13.00x slower by median. |
| Apache POI | ran/lost | `apache-poi` | 82.584 | 176.589 | 0.444 | 352.0 MiB heap | Materialized workbook value read. Tail is noisy; 25.46x slower by median. |
| Polars xlsx2csv | ran/lost | `polars-xlsx2csv` | 114.935 | 174.521 | 0.210 | 75.1 MiB | Value read through xlsx2csv path. 35.43x slower by median. |
| Excelize | ran/lost | `excelize` | 119.388 | 124.169 | 0.026 | 47.3 MiB | File-path materialized value read. Lower memory; 36.80x slower by median. |
| Polars openpyxl | ran/lost | `polars-openpyxl` | 219.728 | 247.138 | 0.148 | 118.8 MiB | Value read through openpyxl engine. 67.73x slower by median. |
| pyopenxlsx | ran/lost | `pyopenxlsx` | 495.510 | 689.964 | 0.286 | 74.7 MiB | Cell materialization. Tail is noisy; 152.75x slower by median. |
| ClosedXML | blocked | `closedxml` | n/a | n/a | n/a | n/a | Not counted. The .NET runner failed with `CSSM_ModuleLoad(): One or more parameters passed to a function were not valid.` |
| fastxlsx | runner unavailable | `fastxlsx` | n/a | n/a | n/a | n/a | Not counted. Python runner failed because `fastxlsx` is not installed in the clean Python environment. |

In-process JavaScript reference:

| Library | Status | Median ms | P95 ms | CV | Peak RSS |
| --- | --- | ---: | ---: | ---: | ---: |
| Ascend in-process | ran/won | 2.911 | 3.259 | 0.081 | 190.6 MiB |
| SheetJS `xlsx@0.18.5` | ran/lost | 26.839 | 30.851 | 0.067 | 285.3 MiB |
| ExcelJS `4.4.0` | ran/lost | 42.433 | 55.173 | 0.151 | 317.6 MiB |

Coverage gate result: failed only because the single-workload scoreboard cannot satisfy earlier workload rows and ClosedXML remains blocked on `warm-workflow`. The warm-workflow comparable rows are ran/won for Ascend against completed comparable runners, but broad speed wording still requires a single clean full-profile promotion run or explicit per-runner blockers.

Humble allowed wording:

> On the generated `warm-workflow` 2000x20 raw OOXML workload at commit `add13c79`, Ascend's warm value-read operation rows were faster than the completed comparable external runners. ClosedXML and fastxlsx are still blocker/unavailable rows and are not counted.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend beats ClosedXML or fastxlsx" from this run.
- Any wording that counts blocked/unavailable runners or a single-workload run as a full-profile win.

Next action: defer production optimization. Run or assemble a current-commit full-profile gate next; if it still fails only for explicit blocked, unavailable, unsupported, or semantic-mismatch rows, keep the claim scoped and attack the highest-impact blocker.

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

Defer: yes. Defer all broad read-speed wording until a clean-tree full `xlsx-read-sota` run either passes coverage or records per-runner blockers without counting them as wins. The immediate next action is a current-commit full-profile gate or merged profile artifact, plus runner hardening for ClosedXML, fastxlsx, rich-metadata semantic mismatches, selected-sheet unsupported-operation gaps, and metadata-only unsupported-operation gaps.
