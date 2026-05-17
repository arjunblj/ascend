# Performance Claim Baseline Matrix

Updated: 2026-05-17

## Current Claim Gate

Status: defer.

This document is the tracked release claim artifact. `/private/tmp` paths below
are reproducibility pointers to the clean run outputs, not required context for
release wording.

No broad XLSX read, XLSX write, SOTA, or QSS-leapfrog speed claim is promotable from this artifact. The cycles below are useful external baseline evidence for scoped release workflows, but they are not a clean full-profile claim because:

- `competitive-scoreboard --require-profile xlsx-read-sota` still fails for broad promotion. The current full-profile run at `9ddfff91` reports no leader failures, and a merged selected-sheet/metadata-only scoreboard removes those same-lane comparability failures, but ClosedXML coverage, feature-rich semantic mismatches, and unsupported selected-sheet/metadata-only competitors remain explicit blockers.
- `competitive-scoreboard --require-profile xlsx-write-sota` is not promotable
  from the current artifact. The completed write rows are scoped wins or
  not-comparable boundaries, and a broad all-workload/all-runner generated-write
  gate at `eca32509` was killed as an operational blocker after producing no
  JSON. Treat broad write wording as downgraded until the gate is split into
  attributable workload groups or produces complete coverage.
- Current focused `plain-text` and `string-heavy` write coverage proves
  ClosedXML and NPOI now run and pass validation on those generated value-write
  rows. Older ClosedXML `CSSM_ModuleLoad()`/unavailable wording is historical
  for prior clean runs, not current evidence for the focused rows. Broad write
  wording is still blocked by missing multi-workload/full-profile coverage.
- Current focused TS/JS/Rust `dense-values` write coverage proves Ascend's
  generated writer is faster by median and p95 than SheetJS, ExcelJS, and
  rust_xlsxwriter on that value-write row. Treat it as scoped row evidence, not
  a broad write-speed or smallest-file claim.
- Current focused TS/JS/Rust `sparse-wide` write coverage supersedes the older
  sparse-wide p95 boundary for current wording: Ascend is faster by median and
  p95 than SheetJS, ExcelJS, and rust_xlsxwriter on the current row, while
  rust_xlsxwriter still uses less RSS and emits a smaller file.
- Current focused TS/JS/Rust `plain-text` write coverage proves Ascend's
  generated writer is faster by median and p95 than SheetJS, ExcelJS, and
  rust_xlsxwriter on that value-write row. Treat it as scoped row evidence, not
  a broad write-speed, lowest-memory, or smallest-file claim.
- Current focused TS/JS/Rust `string-heavy` write coverage proves Ascend's
  generated writer is faster by median and p95 than SheetJS, ExcelJS, and
  rust_xlsxwriter on that value-write row. This supersedes the older noisy
  string-heavy tail boundary for current TS/JS/Rust wording, but not for
  lowest-memory or smallest-file wording.
- Current focused TS/JS/Rust `styles-heavy` write coverage proves Ascend's
  generated writer is faster by median and p95 than SheetJS, ExcelJS, and
  rust_xlsxwriter on that value-write row. Treat it as scoped value evidence,
  not style-fidelity equivalence, lowest-memory, or smallest-file evidence.
- Current focused TS/JS/Rust `formula-heavy` write coverage now includes
  formula-capable SheetJS and ExcelJS rows. Ascend is faster by median and p95
  than SheetJS, ExcelJS, and rust_xlsxwriter on that formula-write row, but
  rust_xlsxwriter uses less RSS and ExcelJS/rust_xlsxwriter emit smaller files.
- The recorded cycles cover public/reproducible generated `dense-values`, `sparse-wide`, `styles-heavy`, `formula-heavy`, `table-heavy`, `feature-rich`, `selected-sheet`, `metadata-only`, `warm-workflow`, and `string-heavy` workloads over `raw-ooxml`, but they are per-workload evidence rows rather than one clean all-workload promotion run.
- Current harness evidence now supports same-lane selected-sheet rows for Ascend, SheetJS, OpenPyXL, and python-calamine. Treat older `openpyxl` and Calamine selected-sheet `unsupported-operation` wording as historical for the recorded clean runs.
- Current harness evidence now supports same-lane metadata-only rows for Ascend, SheetJS, OpenPyXL, and python-calamine. Calamine wins that head-to-head; treat older metadata-only `missing-comparable` or Calamine `unsupported-operation` wording as historical.
- Current `36b927f9` metadata-only recheck still has python-calamine as the
  median and p95 winner for the comparable plain sheet-list/no-cell-hydration
  contract. Ascend beats SheetJS and OpenPyXL on that row, but the metadata-only
  speed claim remains downgraded against the Rust floor.
- Current harness evidence now supports a SheetJS feature-rich rich-metadata row using SheetJS `bookFiles`; older SheetJS `semantic-mismatch` wording is historical for the pre-runner-fix cycles. Calamine-family rich-metadata rows remain not comparable.
- Current focused JS `feature-rich` write evidence is a quality boundary:
  SheetJS is explicitly unsupported for the tracked rich-metadata write
  contract, and ExcelJS runs but is semantically ineligible because it misses a
  tracked comment obligation. Do not count that as a speed win over JS writers.
- Current formula/calc evidence includes focused HyperFormula indexed
  `INDEX/MATCH`, indexed dirty-key/dirty-value edits, prefix-range full-calc
  `SUM`, and prefix-range dirty-head/dirty-tail rows. They are useful
  formula-engine performance evidence, but they are not XLSX behavior parity,
  Excel compatibility, or broad formula SOTA evidence.
- Current real-workbook evidence includes a tracked `strings_links.xlsx`
  open/inspect boundary. Ascend's SDK value-open surface is correct and faster
  than OpenPyXL/POI/ClosedXML on that row, but Rust Calamine is faster on its
  materialization lane and several lower-level or alternate readers are
  semantic mismatches. Do not promote a broad real-workbook speed claim from it.
- Several external runners were unavailable or blocked in the clean benchmark worktree. They are recorded as blockers, not wins.
- Several timing lanes are semantically related but not one unified timing boundary. Do not collapse in-process, preloaded-bytes, file-path, row-stream, and materialized-workbook timings into a single "wins everything" claim.

Humble allowed wording:

> On the generated `string-heavy` 2000x20 raw OOXML workload, using `fixtures/benchmarks/competitive-io.ts` with 5 samples and 1 warmup from a clean detached worktree, Ascend's value-read paths were competitive with or faster than the external readers that successfully ran. Several external runners were unavailable or blocked, so this is one scoped baseline row, not a full `xlsx-read-sota` claim.

Forbidden wording:

- "Ascend is the fastest XLSX reader."
- "Ascend is SOTA for XLSX read."
- "Ascend is SOTA for XLSX write."
- "Ascend beats every external library."
- Any wording that treats failed or unavailable runners as wins.

Next action: downgrade the broad speed claim and stop production optimization from winning rows. Continue only if the performance loop is explicitly attacking a remaining claim blocker or measured loss: ClosedXML coverage, feature-rich semantic mismatches for SheetJS/Calamine, metadata-only versus Calamine, remaining unsupported selected-sheet/metadata-only competitors, or FastXLSX environment coverage.

## Cycle: Formula SOTA Indexed Lookup HyperFormula Row at `cd1c0415`

Classification: comparable formula-engine evidence plus defer. Ascend is the
median and p95 winner on this focused indexed `INDEX/MATCH` formula workflow,
so no production optimization is justified from the row. This is not a broad
formula/calc SOTA claim.

Workflow: exact `INDEX(C$1:C$8000,MATCH(E<n>,A$1:A$8000,0))` formulas over an
8,000-row keyed table, comparing Ascend lookup caching against HyperFormula
with `useColumnIndex` enabled.

Why it matters for release: formula/calc performance is a named performance
lane after read/write baselines. HyperFormula is the strongest direct OSS
formula-engine baseline in this repo, but this row measures generated in-memory
formula calculation only; it does not prove Excel-compatible formula behavior,
XLSX preservation, cached-value truth, or real-workbook parity.

Public/tracked-clean input: `formula-sota` generated the `hf-indexed-index-match`
workload from tracked benchmark code in a clean detached worktree at commit
`cd1c0415`. No private corpus or local research workbook was used. The row used
8,000 data rows, 1,000 formulas, 30 measured samples, and 5 warmups.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-formula-sota-current-cd1c0415 cd1c0415bf3137d89d2b79ad5b64540c1065a434
cd /private/tmp/ascend-formula-sota-current-cd1c0415
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-formula-sota-current-cd1c0415-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/formula-sota.ts --profile hf-indexed-index-match --repeat 30 --warmup 5 --assert-correctness --json > /private/tmp/ascend-formula-sota-current-cd1c0415-runs/hf-indexed-index-match-repeat30.json 2> /private/tmp/ascend-formula-sota-current-cd1c0415-runs/hf-indexed-index-match-repeat30-time.txt
```

Environment:

- Commit: `cd1c0415bf3137d89d2b79ad5b64540c1065a434`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-formula-sota-current-cd1c0415`; `git status --short
  --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `22.22.0`
- HyperFormula dependency: `^3.2.0`
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runtime profile: `profile hf-indexed-index-match`, `rows 8000`, `formulas
  1000`, `repeat 30`, `warmup 5`, `assertCorrectness true`.

Raw output:

```text
/private/tmp/ascend-formula-sota-current-cd1c0415-runs/hf-indexed-index-match-repeat30.json
/private/tmp/ascend-formula-sota-current-cd1c0415-runs/hf-indexed-index-match-repeat30-time.txt
```

Focused formula-engine row, repeat 30 after 5 warmups:

| Engine | Status | Setup median / p95 / CV | Operation median / p95 / CV | Total median / p95 / CV | Correctness | Memory |
| --- | --- | ---: | ---: | ---: | --- | --- |
| Ascend | ran/won | 6.305 ms / 12.225 ms / 0.398 | 29.644 ms / 58.706 ms / 0.379 | 36.690 ms / 67.769 ms / 0.341 | 1,000 changed cells, 0 errors, probe value matched expected `21259` | Process-level peak RSS shared by both engines: 394.2 MiB maximum resident set size; 284.4 MiB peak memory footprint. |
| HyperFormula | ran/lost vs Ascend | 1075.377 ms / 1736.006 ms / 0.295 | 437.929 ms / 1031.233 ms / 0.471 | 1539.467 ms / 2709.070 ms / 0.321 | Probe value matched expected `21259`; `useColumnIndex` enabled | Process-level peak RSS shared by both engines: 394.2 MiB maximum resident set size; 284.4 MiB peak memory footprint. |

Comparison: `operationSpeedupVsHyperFormula: 14.773x`;
`totalSpeedupVsHyperFormula: 41.959x`. Operation sample ranges were
`17.894..67.884 ms` for Ascend and `266.366..1048.998 ms` for HyperFormula.

Rejected broad command: `--profile all --repeat 15 --warmup 3
--assert-correctness --json` was killed after more than two minutes with a
zero-byte JSON file. It is not evidence for any formula speed claim.

Semantic boundary: both engines calculate the same generated in-memory formula
shape and pass the same probe-value assertion. HyperFormula is a formula engine
baseline, not an XLSX reader/writer/preservation engine. This row does not cover
Excel/LibreOffice oracle behavior, formula coverage breadth, cached formula
values, dependency edits after workbook operations, or XLSX roundtrip fidelity.

Humble allowed wording:

> On the generated `hf-indexed-index-match` formula-engine workflow at commit
> `cd1c0415`, Ascend was faster by median and p95 than HyperFormula `^3.2.0`
> with `useColumnIndex` enabled. This is scoped generated formula calculation
> evidence, not broad formula parity or XLSX behavior evidence.

Forbidden wording:

- "Ascend is SOTA for formula calculation."
- "Ascend beats HyperFormula on every formula workflow."
- "Ascend proves Excel-compatible formula parity."
- "Ascend proves cached formula truth or workbook formula preservation."
- Any wording that treats the killed all-profile command as evidence.

Next action: defer production optimization from this winning row. Continue
formula/calc performance work only with a named HyperFormula workflow that
loses, an external formula oracle boundary, or a smaller attributable all-profile
gate that emits complete JSON.

## Cycle: Formula SOTA Indexed Lookup Dirty Edits at `2700c72a`

Classification: comparable formula-engine evidence plus defer. Ascend is the
median and p95 winner on both focused indexed `INDEX/MATCH` incremental edit
workflows. No production optimization is justified from these rows.

Workflow: after initial calculation of exact
`INDEX(C$1:C$8000,MATCH(E<n>,A$1:A$8000,0))` formulas over an 8,000-row keyed
table, measure two incremental edits: changing one lookup key and changing one
indexed return value.

Why it matters for release: these rows exercise dependency-update performance
for lookup-heavy models after an agent edits either the lookup key or the
matched return value. HyperFormula runs with `useColumnIndex` enabled, making it
the relevant OSS formula-engine baseline for this shape. These are generated
in-memory formula-engine rows only; they do not prove Excel-compatible formula
coverage, cached formula truth, dependency updates after XLSX structural edits,
or XLSX roundtrip behavior.

Public/tracked-clean input: `formula-sota` generated the
`hf-indexed-index-match-dirty-key` and `hf-indexed-index-match-dirty-value`
workloads from tracked benchmark code in a clean detached worktree at commit
`2700c72a`. No private corpus or local research workbook was used. Both rows
used 8,000 data rows, 1,000 formulas, 30 measured samples, and 5 warmups.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-formula-indexed-dirty-current-2700c72a 2700c72a8d844fcab6f080547fafe7fbce831b5e
cd /private/tmp/ascend-formula-indexed-dirty-current-2700c72a
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/formula-sota.ts --profile hf-indexed-index-match-dirty-key --repeat 30 --warmup 5 --assert-correctness --json > /private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-key-repeat30.json 2> /private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-key-repeat30-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/formula-sota.ts --profile hf-indexed-index-match-dirty-value --repeat 30 --warmup 5 --assert-correctness --json > /private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-value-repeat30.json 2> /private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-value-repeat30-time.txt
```

Environment:

- Commit: `2700c72a8d844fcab6f080547fafe7fbce831b5e`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-formula-indexed-dirty-current-2700c72a`; `git status
  --short --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `22.22.0`
- HyperFormula dependency: `^3.2.0`
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runtime profile: `rows 8000`, `formulas 1000`, `repeat 30`, `warmup 5`,
  `assertCorrectness true`.

Raw output:

```text
/private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-key-repeat30.json
/private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-key-repeat30-time.txt
/private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-value-repeat30.json
/private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-value-repeat30-time.txt
```

Focused formula-engine rows, repeat 30 after 5 warmups:

| Workflow | Engine | Status | Setup median / p95 / CV | Operation median / p95 / CV | Total median / p95 / CV | Correctness | Memory |
| --- | --- | --- | ---: | ---: | ---: | --- | --- |
| Dirty key edit | Ascend | ran/won | 17.453 ms / 23.353 ms / 0.206 | 0.152 ms / 0.552 ms / 1.406 | 17.591 ms / 23.494 ms / 0.212 | Edited `Sheet1!E1000`, changed 1 cell, 0 errors, probe value matched expected `49259`. | Process-level peak RSS shared by both engines: 458.6 MiB maximum resident set size; 317.9 MiB peak memory footprint. |
| Dirty key edit | HyperFormula | ran/lost vs Ascend | 1252.808 ms / 2396.354 ms / 0.401 | 0.659 ms / 2.049 ms / 0.631 | 1254.083 ms / 2397.869 ms / 0.401 | Edit changed 2 cells; probe value matched expected `49259`. | Process-level peak RSS shared by both engines: 458.6 MiB maximum resident set size; 317.9 MiB peak memory footprint. |
| Dirty return-value edit | Ascend | ran/won | 17.903 ms / 27.906 ms / 0.240 | 0.098 ms / 0.221 ms / 1.761 | 18.008 ms / 28.001 ms / 0.237 | Edited `Sheet1!C3037`, changed 1 cell, 0 errors, probe value matched expected `9000999`. | Process-level peak RSS shared by both engines: 417.9 MiB maximum resident set size; 276.9 MiB peak memory footprint. |
| Dirty return-value edit | HyperFormula | ran/lost vs Ascend | 1463.408 ms / 3140.009 ms / 0.373 | 347.289 ms / 951.250 ms / 0.448 | 1834.094 ms / 3740.465 ms / 0.354 | Edit changed 2 cells; probe value matched expected `9000999`. | Process-level peak RSS shared by both engines: 417.9 MiB maximum resident set size; 276.9 MiB peak memory footprint. |

Comparisons:

- Dirty key edit: `operationSpeedupVsHyperFormula: 4.335x`;
  `totalSpeedupVsHyperFormula: 71.290x`.
- Dirty return-value edit: `operationSpeedupVsHyperFormula: 3532.498x`;
  `totalSpeedupVsHyperFormula: 101.851x`.

Semantic boundary: both engines calculate the same generated in-memory edits
and pass the same probe-value assertions. Changed-cell counts differ because
Ascend reports one changed output cell while HyperFormula reports two changed
cells from its edit API. Ascend operation samples are noisy at sub-millisecond
scale (`opCv 1.406` for dirty-key and `1.761` for dirty-value), so wording
should prefer median/p95 row evidence over microsecond-scale precision. These
rows are formula-engine timing evidence, not XLSX behavior, formula-corpus
parity, or Excel/LibreOffice oracle evidence.

Humble allowed wording:

> On the generated indexed `INDEX/MATCH` dirty-key and dirty-return-value
> workflows at commit `2700c72a`, Ascend had faster median and p95 incremental
> recalculation than HyperFormula `^3.2.0` with `useColumnIndex` enabled. This
> is scoped generated formula-engine evidence, not broad formula parity or XLSX
> behavior evidence.

Forbidden wording:

- "Ascend is SOTA for formula calculation."
- "Ascend beats HyperFormula on every lookup or incremental recalculation
  workflow."
- "Ascend proves Excel-compatible formula parity."
- "Ascend proves workbook formula preservation or cached formula truth."
- Any wording that hides the noisy sub-millisecond Ascend operation samples or
  changed-cell reporting differences.

Next action: defer production optimization from these winning rows. Continue
formula/calc performance work only with a named HyperFormula workflow that
loses, an external formula oracle boundary, or an attributable all-profile gate
that emits complete JSON.

## Cycle: Formula SOTA Prefix Full-Calc HyperFormula Row at `0a9c2b80`

Classification: comparable formula-engine evidence plus defer. Ascend is the
median and p95 winner on this focused full-calculation prefix range workflow.
No production optimization is justified from the row.

Workflow: initial calculation of 5,000 growing `SUM(A$1:A<n>)` formulas over
5,000 source rows, comparing Ascend's formula engine with HyperFormula on the
documented optimized range-composition shape.

Why it matters for release: this is the base full-calc version of the
HyperFormula dependency-graph range composition example. It is a direct JS
formula-engine comparison for a prefix aggregation shape, but it is generated
in-memory formula-engine evidence only; it does not prove Excel-compatible
formula coverage, cached formula truth, or XLSX roundtrip behavior.

Public/tracked-clean input: `formula-sota` generated the
`hf-prefix-range-sum` workload from tracked benchmark code in a clean detached
worktree at commit `0a9c2b80`. No private corpus or local research workbook was
used. The row used 5,000 source rows, 5,000 formulas, `SUM`, 30 measured
samples, and 5 warmups.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-formula-prefix-current-0a9c2b80 0a9c2b801f509b96c4733bbd1f7d91945a0090ee
cd /private/tmp/ascend-formula-prefix-current-0a9c2b80
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-formula-prefix-current-0a9c2b80-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/formula-sota.ts --profile hf-prefix-range-sum --aggregate SUM --repeat 30 --warmup 5 --assert-correctness --json > /private/tmp/ascend-formula-prefix-current-0a9c2b80-runs/hf-prefix-range-sum-repeat30.json 2> /private/tmp/ascend-formula-prefix-current-0a9c2b80-runs/hf-prefix-range-sum-repeat30-time.txt
```

Environment:

- Commit: `0a9c2b801f509b96c4733bbd1f7d91945a0090ee`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-formula-prefix-current-0a9c2b80`; `git status --short
  --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `22.22.0`
- HyperFormula dependency: `^3.2.0`
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runtime profile: `profile hf-prefix-range-sum`, `aggregate SUM`, `rows
  5000`, `formulas 5000`, `repeat 30`, `warmup 5`, `assertCorrectness true`.

Raw output:

```text
/private/tmp/ascend-formula-prefix-current-0a9c2b80-runs/hf-prefix-range-sum-repeat30.json
/private/tmp/ascend-formula-prefix-current-0a9c2b80-runs/hf-prefix-range-sum-repeat30-time.txt
```

Focused formula-engine row, repeat 30 after 5 warmups:

| Engine | Status | Setup median / p95 / CV | Operation median / p95 / CV | Total median / p95 / CV | Correctness | Memory |
| --- | --- | ---: | ---: | ---: | --- | --- |
| Ascend | ran/won | 0.564 ms / 0.694 ms / 0.091 | 4.185 ms / 5.461 ms / 0.137 | 4.744 ms / 6.070 ms / 0.121 | 5,000 changed cells, 0 errors, last value matched expected `12502500`. | Process-level peak RSS shared by both engines: 239.6 MiB maximum resident set size; 191.7 MiB peak memory footprint. |
| HyperFormula | ran/lost vs Ascend | 18.994 ms / 21.820 ms / 0.063 | 10.610 ms / 11.687 ms / 0.048 | 29.536 ms / 32.812 ms / 0.049 | Last value matched expected `12502500`. | Process-level peak RSS shared by both engines: 239.6 MiB maximum resident set size; 191.7 MiB peak memory footprint. |

Comparison: `operationSpeedupVsHyperFormula: 2.536x`;
`totalSpeedupVsHyperFormula: 6.226x`. Operation sample ranges were
`3.784..5.585 ms` for Ascend and `10.058..11.910 ms` for HyperFormula.

Semantic boundary: both engines calculate the same generated in-memory prefix
`SUM` formulas and pass the same final-value assertion. This row is
formula-engine timing evidence, not XLSX behavior, formula-corpus parity, or
Excel/LibreOffice oracle evidence.

Humble allowed wording:

> On the generated `hf-prefix-range-sum/SUM` workflow at commit `0a9c2b80`,
> Ascend had faster median and p95 full calculation than HyperFormula `^3.2.0`
> on the same generated prefix-range formulas. This is scoped formula-engine
> evidence, not broad formula parity or XLSX behavior evidence.

Forbidden wording:

- "Ascend is SOTA for formula calculation."
- "Ascend beats HyperFormula on every full-calculation workflow."
- "Ascend proves Excel-compatible formula parity."
- "Ascend proves workbook formula preservation or cached formula truth."

Next action: defer production optimization from this winning row. Continue
formula/calc performance work only with a named HyperFormula workflow that
loses, an external formula oracle boundary, or an attributable all-profile gate
that emits complete JSON.

## Cycle: Formula SOTA Prefix Dirty-Tail HyperFormula Row at `c06bba18`

Classification: comparable formula-engine evidence plus defer. Ascend is the
median and p95 winner on this focused incremental dirty recalculation workflow.
No production optimization is justified from the row.

Workflow: after an initial full recalc over 5,000 growing
`SUM(A$1:A<n>)` formulas, edit the last source row and measure dirty
propagation through the affected prefix formula.

Why it matters for release: incremental recalculation is a practical calc
workflow after agent edits. This row checks a HyperFormula-documented optimized
growing-range shape, but it is generated in-memory formula-engine evidence only;
it does not prove Excel-compatible formula coverage, cached formula truth, or
XLSX roundtrip behavior.

Public/tracked-clean input: `formula-sota` generated the
`hf-prefix-range-dirty-tail` workload from tracked benchmark code in a clean
detached worktree at commit `c06bba18`. No private corpus or local research
workbook was used. The row used 5,000 source rows, 5,000 formulas, `SUM`, 30
measured samples, and 5 warmups.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-formula-dirty-current-c06bba18 c06bba18
cd /private/tmp/ascend-formula-dirty-current-c06bba18
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-formula-dirty-current-c06bba18-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/formula-sota.ts --profile hf-prefix-range-dirty-tail --aggregate SUM --repeat 30 --warmup 5 --assert-correctness --json > /private/tmp/ascend-formula-dirty-current-c06bba18-runs/hf-prefix-range-dirty-tail-repeat30.json 2> /private/tmp/ascend-formula-dirty-current-c06bba18-runs/hf-prefix-range-dirty-tail-repeat30-time.txt
```

Environment:

- Commit: `c06bba18`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-formula-dirty-current-c06bba18`; `git status --short
  --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `22.22.0`
- HyperFormula dependency: `^3.2.0`
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runtime profile: `profile hf-prefix-range-dirty-tail`, `aggregate SUM`,
  `rows 5000`, `formulas 5000`, `repeat 30`, `warmup 5`,
  `assertCorrectness true`.

Raw output:

```text
/private/tmp/ascend-formula-dirty-current-c06bba18-runs/hf-prefix-range-dirty-tail-repeat30.json
/private/tmp/ascend-formula-dirty-current-c06bba18-runs/hf-prefix-range-dirty-tail-repeat30-time.txt
```

Focused formula-engine row, repeat 30 after 5 warmups:

| Engine | Status | Setup median / p95 / CV | Operation median / p95 / CV | Total median / p95 / CV | Correctness | Memory |
| --- | --- | ---: | ---: | ---: | --- | --- |
| Ascend | ran/won | 7.122 ms / 9.880 ms / 0.185 | 0.059 ms / 0.090 ms / 0.208 | 7.184 ms / 9.941 ms / 0.183 | Tail edit changed 1 cell, 0 errors, probe value matched expected `13497503`. | Process-level peak RSS shared by both engines: 247.8 MiB maximum resident set size; 183.7 MiB peak memory footprint. |
| HyperFormula | ran/lost vs Ascend | 53.358 ms / 65.878 ms / 0.171 | 0.091 ms / 0.516 ms / 1.157 | 53.496 ms / 65.946 ms / 0.172 | Tail edit changed 2 cells; probe value matched expected `13497503`. | Process-level peak RSS shared by both engines: 247.8 MiB maximum resident set size; 183.7 MiB peak memory footprint. |

Comparison: `operationSpeedupVsHyperFormula: 1.531x`;
`totalSpeedupVsHyperFormula: 7.447x`. HyperFormula's operation tail is much
noisier in this row, with `opCv 1.157` and p95 `0.516 ms`; Ascend's p95 is
`0.090 ms`.

Semantic boundary: both engines calculate the same generated in-memory dirty
tail edit and pass the same probe-value assertion. Changed-cell counts differ
because Ascend reports 1 changed output cell while HyperFormula reports 2
changed cells from the edit API. This row is formula-engine timing evidence,
not XLSX behavior, formula-corpus parity, or Excel/LibreOffice oracle evidence.

Humble allowed wording:

> On the generated `hf-prefix-range-dirty-tail/SUM` workflow at commit
> `c06bba18`, Ascend had faster median and p95 dirty recalculation than
> HyperFormula `^3.2.0` on the same generated tail edit. This is scoped
> formula-engine evidence, not broad formula parity or XLSX behavior evidence.

Forbidden wording:

- "Ascend is SOTA for formula calculation."
- "Ascend beats HyperFormula on every incremental recalculation workflow."
- "Ascend proves Excel-compatible formula parity."
- "Ascend proves workbook formula preservation or cached formula truth."

Next action: defer production optimization from this winning row. Continue
formula/calc performance work only with a named HyperFormula workflow that
loses or an attributable all-profile gate that emits complete JSON.

## Cycle: Formula SOTA Prefix Dirty-Head HyperFormula Row at `bd91386a`

Classification: comparable formula-engine evidence plus defer. Ascend is the
median and p95 winner on this focused incremental dirty recalculation workflow.
No production optimization is justified from the row.

Workflow: after an initial full recalc over 5,000 growing
`SUM(A$1:A<n>)` formulas, edit `A1` and measure dirty propagation through all
5,000 affected prefix formulas.

Why it matters for release: this is the head-edit version of HyperFormula's
documented optimized growing-range recalculation shape. It is the more expensive
dirty-prefix case than the tail edit, but it remains generated in-memory
formula-engine evidence only; it does not prove Excel-compatible formula
coverage, cached formula truth, or XLSX roundtrip behavior.

Public/tracked-clean input: `formula-sota` generated the
`hf-prefix-range-dirty-head` workload from tracked benchmark code in a clean
detached worktree at commit `bd91386a`. No private corpus or local research
workbook was used. The row used 5,000 source rows, 5,000 formulas, `SUM`, 30
measured samples, and 5 warmups.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-formula-dirty-head-current-bd91386a bd91386a
cd /private/tmp/ascend-formula-dirty-head-current-bd91386a
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-formula-dirty-head-current-bd91386a-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/formula-sota.ts --profile hf-prefix-range-dirty-head --aggregate SUM --repeat 30 --warmup 5 --assert-correctness --json > /private/tmp/ascend-formula-dirty-head-current-bd91386a-runs/hf-prefix-range-dirty-head-repeat30.json 2> /private/tmp/ascend-formula-dirty-head-current-bd91386a-runs/hf-prefix-range-dirty-head-repeat30-time.txt
```

Environment:

- Commit: `bd91386a`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-formula-dirty-head-current-bd91386a`; `git status
  --short --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `22.22.0`
- HyperFormula dependency: `^3.2.0`
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runtime profile: `profile hf-prefix-range-dirty-head`, `aggregate SUM`,
  `rows 5000`, `formulas 5000`, `repeat 30`, `warmup 5`,
  `assertCorrectness true`.

Raw output:

```text
/private/tmp/ascend-formula-dirty-head-current-bd91386a-runs/hf-prefix-range-dirty-head-repeat30.json
/private/tmp/ascend-formula-dirty-head-current-bd91386a-runs/hf-prefix-range-dirty-head-repeat30-time.txt
```

Focused formula-engine row, repeat 30 after 5 warmups:

| Engine | Status | Setup median / p95 / CV | Operation median / p95 / CV | Total median / p95 / CV | Correctness | Memory |
| --- | --- | ---: | ---: | ---: | --- | --- |
| Ascend | ran/won | 6.284 ms / 8.729 ms / 0.372 | 0.810 ms / 0.990 ms / 0.110 | 7.054 ms / 9.612 ms / 0.338 | Head edit changed 5,000 cells, 0 errors, probe value matched expected `13502502`. | Process-level peak RSS shared by both engines: 253.7 MiB maximum resident set size; 195.8 MiB peak memory footprint. |
| HyperFormula | ran/lost vs Ascend | 39.465 ms / 46.316 ms / 0.085 | 11.719 ms / 13.636 ms / 0.072 | 51.246 ms / 59.952 ms / 0.073 | Head edit changed 5,001 cells; probe value matched expected `13502502`. | Process-level peak RSS shared by both engines: 253.7 MiB maximum resident set size; 195.8 MiB peak memory footprint. |

Comparison: `operationSpeedupVsHyperFormula: 14.474x`;
`totalSpeedupVsHyperFormula: 7.265x`. Ascend's operation p95 was `0.990 ms`;
HyperFormula's operation p95 was `13.636 ms`.

Semantic boundary: both engines calculate the same generated in-memory dirty
head edit and pass the same probe-value assertion. Changed-cell counts differ
because Ascend reports 5,000 changed output cells while HyperFormula reports
5,001 changed cells from the edit API. This row is formula-engine timing
evidence, not XLSX behavior, formula-corpus parity, or Excel/LibreOffice oracle
evidence.

Humble allowed wording:

> On the generated `hf-prefix-range-dirty-head/SUM` workflow at commit
> `bd91386a`, Ascend had faster median and p95 dirty recalculation than
> HyperFormula `^3.2.0` on the same generated head edit. This is scoped
> formula-engine evidence, not broad formula parity or XLSX behavior evidence.

Forbidden wording:

- "Ascend is SOTA for formula calculation."
- "Ascend beats HyperFormula on every incremental recalculation workflow."
- "Ascend proves Excel-compatible formula parity."
- "Ascend proves workbook formula preservation or cached formula truth."

Next action: defer production optimization from this winning row. Continue
formula/calc performance work only with a named HyperFormula workflow that
loses or an attributable all-profile gate that emits complete JSON.

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

## Cycle: Metadata-Only Relationship Recovery Profile at `38cd8ec5`

Classification: kill/defer. A clean current rerun still keeps
`python-calamine-metadata-only` as the winner for the comparable plain
metadata-only sheet-list contract. Profiling named one narrow production cost
center, but the smallest safe candidate did not validate in an Ascend-only
repeat-40 A/B, so no production optimization is carried forward.

Workflow: generated XLSX metadata-only open/inspect, loading workbook and sheet
metadata without hydrating cells.

Why it matters for release: this remains the priority read workflow where
Ascend has a measured comparable loss. Any metadata-only speed wording must not
hide the native Calamine sheet-list baseline.

Public/tracked-clean input: `competitive-io` generated `metadata-only`
`raw-ooxml`, 200 rows x 20 columns, three workbook sheets, 15,347 input bytes,
from tracked benchmark code at commit `38cd8ec5`.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-metadata-profile-38cd8ec5 38cd8ec5
cd /private/tmp/ascend-metadata-profile-38cd8ec5
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-metadata-profile-38cd8ec5-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only,sheetjs-metadata-only,openpyxl-metadata-only,python-calamine-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 15 --warmup 3 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-calamine-head-to-head.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-calamine-head-to-head.json --json --metric medianMs --require-profile xlsx-read-sota --assert-profile-leader ascend > /private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-calamine-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/xlsx-read-phase.ts --json --workload metadata-only --read-source raw-ooxml --rows 200 --cols 20 --phase read --repeat 30 --warmup 5 --validation-mode sample --gc-between-samples > /private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-read-phase.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/profile-bun.ts --mode all-md --label xlsx-read-metadata-only-200x20-read -- bun run fixtures/benchmarks/xlsx-read-phase.ts --input-file /private/tmp/ascend-benchmark-inputs/xlsx-read-phase-metadata-only-raw-ooxml-200x20-efec7201ae64a957.xlsx --rows 200 --cols 20 --workload metadata-only --read-source raw-ooxml --phase read --repeat 30 --warmup 5 --validation-mode sample --gc-between-samples --json > /private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-read-profile.log
```

Rejected optimization A/B:

```bash
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 40 --warmup 5 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-ascend-only-repeat40.json
git apply /private/tmp/ascend-metadata-profile-38cd8ec5-runs/reader-early-return.patch
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 40 --warmup 5 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-profile-38cd8ec5-patched-runs/metadata-ascend-only-repeat40.json
```

Environment:

- Commit: `38cd8ec5`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-metadata-profile-38cd8ec5`
- Bun runtime: `1.3.6`
- Node: `22.22.0`
- Python: `3.9.6`
- Platform: Darwin arm64, macOS `26.4`, kernel `25.4.0`
- Runtime profile: `category read`, `executionScope external-process`,
  `workload metadata-only`, `readSource raw-ooxml`, `validationMode each`,
  `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-calamine-head-to-head.json
/private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-calamine-scoreboard.json
/private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-read-phase.json
/private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-ascend-only-repeat40.json
/private/tmp/ascend-metadata-profile-38cd8ec5-patched-runs/metadata-ascend-only-repeat40.json
```

Result, repeat 15 after 3 warmups:

| Runner | Status | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `python-calamine-metadata-only` | ran/won | 0.102 | 1.692 | 1.705 | 28.7 MiB | Same plain sheet-list/no-cell-hydration contract: three sheets loaded, cells not hydrated. |
| `sheetjs-metadata-only` | ran/lost vs Calamine, ran/won vs Ascend median | 1.586 | 3.780 | 0.512 | 145.8 MiB | Same plain metadata contract using SheetJS `bookSheets`. |
| `ascend-external-metadata-only` | ran/lost vs Calamine and SheetJS, ran/won vs OpenPyXL | 3.558 | 8.432 | 0.594 | 88.8 MiB | Same plain metadata contract plus Ascend's SDK metadata-only inspection surface. |
| `openpyxl-metadata-only` | ran/lost | 6.122 | 13.079 | 0.360 | 49.1 MiB | Same plain metadata contract using read-only OpenPyXL metadata inventory. |

Scoreboard result: `profileLeaderFailures` contains
`winner=python-calamine-metadata-only expected=ascend` for
`read-metadata-only`.

Profile result: `xlsx-read-phase` reported `readXlsxMedianMs: 0.640` with
15,347 input bytes. The Bun CPU profile was dominated by harness/module/GC
costs, but within production code it named workbook relationship recovery and
content-type part scanning as a plausible metadata-only cost center:
`recoverWorkbookRelationships`, `availablePartsForContentType`, `sort`, and
`localeCompare` appeared under `readXlsxArchive`.

Rejected optimization: a current-worktree candidate moved the existing
`recoverWorkbookRelationships` early return ahead of worksheet/chartsheet
content-type scans when workbook relationships were already complete. Focused
correctness tests passed, but the repeat-40 Ascend-only A/B did not validate an
improvement: baseline median `2.779 ms`, p95 `13.191 ms`, CV `1.046`, peak RSS
`81.4 MiB`; patched median `3.017 ms`, p95 `13.135 ms`, CV `0.949`, peak RSS
`83.9 MiB`. The patch was reverted.

Semantic boundary: the Calamine row is comparable only for the generated plain
metadata-only sheet-list/no-cell-hydration contract. It is still not evidence
that Calamine supports Ascend's broader safe-open trust inspection for document
properties, active-content inventory, package risk reporting, or edit-planning
decisions.

Humble allowed wording:

> On the generated plain metadata-only workload at `38cd8ec5`, Calamine is
> faster than Ascend for the comparable sheet-list/no-cell-hydration timing
> lane. A profiled relationship-recovery early-return candidate did not improve
> Ascend in repeat-40 validation, so metadata-only speed wording remains
> downgraded.

Forbidden wording:

- "Ascend is fastest for metadata-only XLSX reads."
- "Ascend beats Calamine on metadata-only open."
- "The relationship-recovery early-return patch improves metadata-only reads."
- "Calamine proves faster safe-open trust inspection than Ascend."
- Any wording that treats unsupported safe-open semantics as wins.

Next action: keep the metadata-only speed claim downgraded. Do not revisit this
row with another reader patch until a profile names a larger production cost
center than complete-relationship recovery scans, or until the benchmark moves
to a broader safe-open metadata contract that Calamine can also satisfy.

## Cycle: Current Metadata-Only Rust Floor Recheck at `36b927f9`

Classification: current claim downgrade. The current clean HEAD still has
`python-calamine-metadata-only` as the median and p95 winner for the comparable
plain metadata-only sheet-list/no-cell-hydration contract. Ascend is ahead of
SheetJS and OpenPyXL on the same external-process timing lane, but this remains
a Rust-floor loss and not an optimization win.

Workflow: generated XLSX metadata-only open/inspect, loading workbook and sheet
metadata without hydrating cells.

Why it matters for release: metadata-only open is the first safe inspection path
for unknown workbooks. The Rust floor matters here because Calamine is the
fastest credible open/inspect baseline for plain workbook sheet metadata. Any
metadata-only speed wording must remain downgraded while this comparable row
loses.

Public/tracked-clean input: `competitive-io` generated `metadata-only`
`raw-ooxml`, 200 rows x 20 columns, three workbook sheets, 15,347 input bytes,
from tracked benchmark code at commit `36b927f9`. No private corpus or local
research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-metadata-rust-current-36b927f9 36b927f9af6e954de08978c3ca93e4f3834ddaa7
cd /private/tmp/ascend-metadata-rust-current-36b927f9
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-metadata-rust-current-36b927f9-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only,sheetjs-metadata-only,openpyxl-metadata-only,python-calamine-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 15 --warmup 3 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head.json 2> /private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head.json --json --metric medianMs --require-profile xlsx-read-sota --assert-profile-leader ascend > /private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head.json --json --metric p95Ms --require-profile xlsx-read-sota --assert-profile-leader ascend > /private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head-p95-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/xlsx-read-phase.ts --json --workload metadata-only --read-source raw-ooxml --rows 200 --cols 20 --phase read --repeat 60 --warmup 10 --validation-mode sample --gc-between-samples > /private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-read-phase-repeat60.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --execution-scope external-process --libraries ascend-external-metadata-only --workload metadata-only --read-source raw-ooxml --repeat 60 --warmup 10 --validation-mode each --runner-manifest fixtures/benchmarks/runners/metadata-only-readers.manifest.json > /private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-ascend-only-repeat60.json
```

Environment:

- Commit: `36b927f9af6e954de08978c3ca93e4f3834ddaa7`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-metadata-rust-current-36b927f9`; `git status --short
  --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `26.0.0`
- Python: `3.13.3`
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runtime profile: `category read`, `executionScope external-process`,
  `workload metadata-only`, `readSource raw-ooxml`, `validationMode each`,
  `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head.json
/private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head-time.txt
/private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head-scoreboard.json
/private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-calamine-head-to-head-p95-scoreboard.json
/private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-read-phase-repeat60.json
/private/tmp/ascend-metadata-rust-current-36b927f9-runs/metadata-ascend-only-repeat60.json
```

Current Rust-floor head-to-head, repeat 15 after 3 warmups:

| Runner | Status | Median ms | P95 ms | CV | Peak RSS | Semantic comparability |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `python-calamine-metadata-only` | ran/won median and p95 | 0.068 | 0.212 | 0.625 | 28.3 MiB | Same plain sheet-list/no-cell-hydration contract: three sheets loaded, cells not hydrated. |
| `ascend-external-metadata-only` | ran/lost vs Calamine, ran/won vs SheetJS/openpyxl | 0.342 | 0.579 | 0.240 | 91.3 MiB | Same plain metadata contract plus Ascend's SDK metadata-only inspection surface. |
| `sheetjs-metadata-only` | ran/lost | 0.584 | 0.875 | 0.173 | 152.1 MiB | Same plain metadata contract using SheetJS `bookSheets`. |
| `openpyxl-metadata-only` | ran/lost | 2.032 | 3.525 | 0.252 | 53.0 MiB | Same plain metadata contract using read-only OpenPyXL metadata inventory. |

Process-level `/usr/bin/time -l`: `1.58 real`, `1.10 user`, `0.28 sys`,
`160972800` maximum resident set size, `123306824` peak memory footprint.

Scoreboard result:

- Median scoreboard: metadata-only group winner was
  `python-calamine-metadata-only`; `profileLeaderFailures` contains
  `winner=python-calamine-metadata-only expected=ascend`.
- P95 scoreboard: metadata-only group winner was
  `python-calamine-metadata-only`; `profileLeaderFailures` contains
  `winner=python-calamine-metadata-only expected=ascend`.
- Both scoreboards still have 59 full-profile coverage failures and 7 coverage
  gaps because this is a focused row, not a full `xlsx-read-sota` promotion run.

Phase/profile result: focused `xlsx-read-phase` with sample validation and GC
reported `readXlsxMedianMs: 0.274`, 15,347 input bytes, and 119.9 MiB peak RSS.
Ascend-only external repeat-60 reported median `0.391 ms`, p95 `0.589 ms`, CV
`0.590`, and 94.2 MiB peak RSS. A validation-free CPU-profile run reported
`readXlsxMedianMs: 0.063`; most profile time was harness/module loading rather
than a larger safe production hotspot. The remaining release-facing gap is
therefore mostly the broader SDK/assertion surface and safety inventory, not a
clear isolated reader loop to remove without narrowing semantics.

Semantic boundary: this row is comparable only for the generated plain
metadata-only sheet-list/no-cell-hydration contract. It is not evidence that
Calamine supports Ascend's broader safe-open trust inspection for document
properties, active-content inventory, package risk reporting, or edit-planning
decisions. Conversely, Ascend beating SheetJS and OpenPyXL here is not a
metadata-only SOTA claim because the Rust floor still wins.

Humble allowed wording:

> On the generated plain metadata-only workload at `36b927f9`, Ascend was faster
> than SheetJS and OpenPyXL on the same external-process metadata-only timing
> lane, but python-calamine remained faster by median and p95 for the comparable
> sheet-list/no-cell-hydration contract. Metadata-only speed wording remains
> downgraded against the Rust floor.

Forbidden wording:

- "Ascend is fastest for metadata-only XLSX reads."
- "Ascend beats Calamine on metadata-only open."
- "Ascend has closed the metadata-only Rust-floor gap."
- "Ascend beating SheetJS/OpenPyXL proves a metadata-only SOTA claim."
- "Calamine proves faster safe-open trust inspection than Ascend."

Next action: keep the metadata-only speed claim downgraded. Do not trade away
document-property, active-content, or package-risk metadata to win this plain
sheet-list row. Revisit production optimization only with a profile that names
a larger safe reader hot spot or with a broader Calamine-comparable safe-open
contract.

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

## Cycle: Dense Values Write Current Fastest Comparable Row at `d7b17aa6`

Classification: comparable external evidence plus defer. This refreshes the
dense-values write row at current `HEAD`. Ascend is the median and p95 winner
among the completed fastest comparable writers in this row with low noise. No
production optimization is justified from a winning row.

Workflow: generated XLSX write for dense numeric values, 2000 rows x 20 columns.

Why it matters for release: dense value export is the simplest generated
commit/write path after an agent creates or rewrites workbook values. It is the
lowest-level write row in the `xlsx-write-sota` profile and validates that the
dense streaming fallback remains a current release win.

Public/tracked-clean input: `competitive-io` generated the `dense-values`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `d7b17aa6`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-dense-current-d7b17aa6 d7b17aa6c744
cd /private/tmp/ascend-write-dense-current-d7b17aa6
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-dense-current-d7b17aa6-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,fastexcel-java,sheetjs --workload dense-values --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-dense-current-d7b17aa6-runs/write-dense-values-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-dense-current-d7b17aa6-runs/write-dense-values-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-dense-current-d7b17aa6-runs/write-dense-values-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `d7b17aa6c7447b7dc3e6eadfc303e63d5cc22d51`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-dense-current-d7b17aa6`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload dense-values`, `validationMode each`,
  `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-write-dense-current-d7b17aa6-runs/write-dense-values-fastest-repeat15.json
/private/tmp/ascend-write-dense-current-d7b17aa6-runs/write-dense-values-fastest-repeat15-scoreboard.json
```

Focused fastest comparable writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 2.667 | 3.592 | 0.100 | 79.0 MiB | 172259 |
| `fastexcel-java` | ran/lost vs Ascend | 10.011 | 27.631 | 0.408 | 472.0 MiB | 161598 |
| `excelize` | ran/lost vs Ascend | 15.337 | 17.198 | 0.040 | 17.1 MiB | 120015 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 20.924 | 22.041 | 0.027 | 17.9 MiB | 119134 |
| `sheetjs` | ran/lost vs Ascend | 26.855 | 33.489 | 0.095 | 243.2 MiB | 1181431 |

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
speed here but uses more RSS and emits a larger XLSX than Excelize and
rust_xlsxwriter.

Humble allowed wording:

> On the generated 2000 x 20 dense-values write row at commit `d7b17aa6`,
> Ascend's focused external repeat-15 run had the fastest median and p95 among
> the completed fastest comparable writers in this row. This is scoped generated
> dense-value write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every generated XLSX writer."
- "Ascend beats omitted, unsupported, or blocked dense-value writers."
- "Ascend produces the smallest dense-value XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: defer production optimization for dense-values from this current
winning row. Continue with the next priority workflow only if it can produce a
validated optimization, comparable baseline row, or explicit claim downgrade.

## Cycle: Dense Values TS/JS/Rust Write Head-to-Head at `7cc7e2c3`

Classification: comparable external evidence plus defer. This refreshes the
generated dense-value write row against the TS/JS and Rust libraries that matter
most for the current performance bar. Ascend is the median and p95 winner
against SheetJS, ExcelJS, and rust_xlsxwriter in this row. No production
optimization is justified from a winning row.

Workflow: generated XLSX write for dense numeric values, 2000 rows x 20 columns.

Why it matters for release: TS/JS and Rust are the key head-to-head ecosystems
for Ascend's agent-native TypeScript runtime and low-level performance bar. The
older dense row included SheetJS and rust_xlsxwriter but omitted ExcelJS and was
tied to an older commit; this run provides current clean-commit evidence for
the focused TS/JS/Rust comparison.

Public/tracked-clean input: `competitive-io` generated the `dense-values`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `7cc7e2c3`. No private corpus or local research
workbook was used. The generated sheet has one sheet, 2,000 rows, 20 columns,
and 40,000 logical cells.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-js-rust-current-7cc7e2c3 7cc7e2c352014d551392b6b9e52cbbf9deb0cd21
cd /private/tmp/ascend-write-js-rust-current-7cc7e2c3
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-js-rust-current-7cc7e2c3-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,rust-xlsxwriter --workload dense-values --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-js-rust-current-7cc7e2c3-runs/write-dense-values-js-rust-repeat15.json 2> /private/tmp/ascend-write-js-rust-current-7cc7e2c3-runs/write-dense-values-js-rust-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-js-rust-current-7cc7e2c3-runs/write-dense-values-js-rust-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-js-rust-current-7cc7e2c3-runs/write-dense-values-js-rust-repeat15-scoreboard.json
```

Environment:

- Commit: `7cc7e2c352014d551392b6b9e52cbbf9deb0cd21`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-js-rust-current-7cc7e2c3`; `git status --short
  --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload dense-values`, `validationMode each`,
  `repeat 15`, `warmup 3`.
- Runner versions: SheetJS `0.18.5`, ExcelJS `4.4.0`, rust_xlsxwriter `0.1.0`.

Raw output:

```text
/private/tmp/ascend-write-js-rust-current-7cc7e2c3-runs/write-dense-values-js-rust-repeat15.json
/private/tmp/ascend-write-js-rust-current-7cc7e2c3-runs/write-dense-values-js-rust-repeat15-time.txt
/private/tmp/ascend-write-js-rust-current-7cc7e2c3-runs/write-dense-values-js-rust-repeat15-scoreboard.json
```

Focused TS/JS/Rust row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 6.013 | 7.776 | 0.103 | 81.0 MiB | 172259 |
| `sheetjs` | ran/lost vs Ascend | 65.011 | 119.687 | 0.253 | 240.6 MiB | 1181431 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 70.629 | 148.107 | 0.343 | 21.0 MiB | 119134 |
| `exceljs` | ran/lost vs Ascend | 124.324 | 188.872 | 0.169 | 244.2 MiB | 121315 |

Process-level `/usr/bin/time -l` for the full command reported `56.65 real`,
273,678,336 bytes maximum resident set size, and 112,116,552 bytes peak memory
footprint.

Scoreboard result:

- Focused repeat-15 TS/JS/Rust row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- Full `xlsx-write-sota` coverage still fails, with 59 coverage failures,
  because this row intentionally covers only `dense-values` and only Ascend,
  SheetJS, ExcelJS, and rust_xlsxwriter. Missing workloads and omitted writers
  are not wins.

Semantic comparability: all four rows reopened successfully, matched one sheet
and 40,000 cells, and passed semantic cell value validation. Ascend, SheetJS,
and ExcelJS matched ordered semantic cell hashes. rust_xlsxwriter matched
sorted semantic values but not ordered semantic hashes, so its row is useful
value-write timing evidence but not byte/order-equivalent output evidence.
Ascend wins median and p95 here, but rust_xlsxwriter and ExcelJS emit smaller
XLSX files and rust_xlsxwriter uses less RSS.

Humble allowed wording:

> On the generated 2000 x 20 `dense-values` write row at commit `7cc7e2c3`,
> Ascend's focused external repeat-15 run was faster by median and p95 than
> SheetJS `0.18.5`, ExcelJS `4.4.0`, and rust_xlsxwriter `0.1.0`, with all rows
> passing semantic value validation. This is scoped TS/JS/Rust generated-write
> evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every TS/JS or Rust writer on every workload."
- "Ascend proves byte/order-equivalent output against rust_xlsxwriter."
- "Ascend produces the smallest dense-value XLSX."
- "Ascend uses less memory than rust_xlsxwriter on dense-value writes."

Next action: keep TS/JS/Rust benchmark work focused on rows where Ascend lacks
current evidence or loses. Do not re-optimize this dense-value row unless a
current JS/Rust rerun regresses it.

## Cycle: Sparse Wide Write Current Split Row at `ddb0eb77`

Classification: comparable external evidence plus blocked broad-row boundary.
The all-runner sparse-wide generated-write row was killed as an opaque
zero-byte blocker, then a focused fastest/profile-critical split row completed.
Ascend is the median and p95 winner among the completed focused comparable
writers, but the Ascend and rust_xlsxwriter rows are noisy. This is scoped
sparse-wide write evidence, not a full write-values or `xlsx-write-sota` claim.

Workflow: generated XLSX write for sparse-wide values, 5000 rows x 256 columns.

Why it matters for release: sparse-wide workbooks model dashboards, exported
planning sheets, and agent-generated grids where the logical sheet is wide but
only a fraction of cells are populated. This is one of the value-write workloads
required by the existing `xlsx-write-sota` profile.

Public/tracked-clean input: `competitive-io` generated the `sparse-wide`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `ddb0eb77`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-sparse-current-ddb0eb77 ddb0eb77
cd /private/tmp/ascend-write-sparse-current-ddb0eb77
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-sparse-current-ddb0eb77-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,xlsxwriter,xlsxwriter-constant-memory,pyexcelerate,pyexcelerate-range,pyexcelerate-cell,fastexcel-java,openpyxl,openpyxl-write-only,apache-poi,closedxml,rust-xlsxwriter,excelize --workload sparse-wide --repeat 5 --warmup 1 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-sparse-current-ddb0eb77-runs/write-sparse-wide-repeat5.json
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,fastexcel-java,sheetjs,xlsxwriter --workload sparse-wide --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-sparse-current-ddb0eb77-runs/write-sparse-wide-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-sparse-current-ddb0eb77-runs/write-sparse-wide-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-sparse-current-ddb0eb77-runs/write-sparse-wide-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `ddb0eb772c8be530dda581e7c6aed349bb66008d`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-sparse-current-ddb0eb77`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload sparse-wide`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-sparse-current-ddb0eb77-runs/write-sparse-wide-repeat5.json
/private/tmp/ascend-write-sparse-current-ddb0eb77-runs/write-sparse-wide-fastest-repeat15.json
/private/tmp/ascend-write-sparse-current-ddb0eb77-runs/write-sparse-wide-fastest-repeat15-scoreboard.json
```

Blocked broad row:

| Artifact | Status | Sample count | Median | P95 | CV/noise | Memory/size |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `write-sparse-wide-repeat5.json` | blocked, killed before JSON emission | 0 | n/a | n/a | n/a | n/a |

Focused fastest/profile-critical split row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 30.063 | 102.333 | 0.615 | 171.1 MiB | 228209 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 67.606 | 369.384 | 0.840 | 52.8 MiB | 175581 |
| `fastexcel-java` | ran/lost vs Ascend | 116.010 | 322.587 | 0.505 | 624.0 MiB | 183160 |
| `excelize` | ran/lost vs Ascend | 141.917 | 195.216 | 0.317 | 20.8 MiB | 168716 |
| `sheetjs` | ran/lost vs Ascend | 1029.037 | 2224.981 | 0.341 | 201.3 MiB | 883673 |
| `xlsxwriter` | ran/lost vs Ascend | 1160.794 | 1870.990 | 0.245 | 75.5 MiB | 180517 |

Scoreboard result:

- Focused repeat-15 sparse-wide split row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The scoreboard command exits nonzero for full-profile coverage because this is
  a single split row. Omitted writers from the blocked broad row are not counted
  as Ascend wins.

Semantic comparability: all listed focused rows reopened successfully, matched
the expected sheet and cell-count shape, and passed the benchmark's write
correctness gate. The semantic hash assertions differ by writer ordering lane,
so this row supports value-write comparability, not byte/order-equivalent output
claims. Ascend is not the smallest output and uses more RSS than Excelize,
rust_xlsxwriter, and XlsxWriter.

Humble allowed wording:

> On the generated sparse-wide write row at commit `ddb0eb77`, Ascend's focused
> external repeat-15 split run had the fastest median and p95 among the completed
> focused comparable writers. The broader all-runner sparse-wide row was blocked
> before JSON emission, so this is scoped sparse-wide write evidence, not a
> broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every sparse-wide writer."
- "Ascend has a low-noise sparse-wide tail win."
- "Ascend beats omitted, unsupported, blocked, or untested sparse-wide writers."
- "Ascend produces the smallest sparse-wide XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: defer production optimization for sparse-wide from this focused
median/p95 win, keep the broad sparse-wide row blocked, and continue splitting
the write-values profile into attributable rows.

## Cycle: Sparse Wide Write Current Tail Boundary at `6595d42c`

Classification: comparable external evidence plus p95/tail boundary. Ascend is
the median winner in the current fastest-comparable sparse-wide generated write
row, but `rust-xlsxwriter` is the p95 winner and uses less RSS with a smaller
output file. This is a scoped median win and a tail-latency claim downgrade, not
a broad `xlsx-write-sota` claim or a production optimization target.

Workflow: generated XLSX write for sparse-wide values, 5000 rows x 256 columns,
23,093 populated cells.

Why it matters for release: sparse-wide exports are a required write-values row
in the `xlsx-write-sota` profile. Broad write-speed wording must cover both the
median and tail behavior of sparse agent-generated workbooks instead of
promoting median-only wins.

Public/tracked-clean input: `competitive-io` generated the `sparse-wide`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `6595d42c`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-values-current-6595d42c 6595d42c
cd /private/tmp/ascend-write-values-current-6595d42c
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-values-current-6595d42c-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,fastexcel-java,sheetjs,xlsxwriter --workload sparse-wide --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15.json --json --metric p95Ms --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15-p95-scoreboard.json
```

Environment:

- Commit: `6595d42c38b2d089a96251755e7a8af86c3e4e6a`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-values-current-6595d42c`
- Benchmark runtime metadata: Bun `1.3.13`, Node `24.3.0`, Darwin arm64
- Shell runtime checks: `/Users/arjun/.bun/bin/bun --version` reported `1.3.6`,
  Python `3.9.6`, Go `1.26.3 darwin/arm64`, macOS `26.4`, kernel `25.4.0`
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload sparse-wide`, `validationMode each`,
  `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15.json
/private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15-scoreboard.json
/private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15-p95-scoreboard.json
```

Focused fastest comparable writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | median winner, p95 loss vs `rust-xlsxwriter` | 27.860 | 173.863 | 0.919 | 117.1 MiB | 228209 |
| `rust-xlsxwriter` | median loss vs Ascend, p95 winner | 41.707 | 54.644 | 0.203 | 53.6 MiB | 175581 |
| `fastexcel-java` | ran/lost vs Ascend median | 68.097 | 493.538 | 1.038 | 568.0 MiB | 183159 |
| `excelize` | ran/lost vs Ascend median | 121.323 | 242.010 | 0.334 | 21.3 MiB | 168716 |
| `xlsxwriter` | ran/lost vs Ascend median | 860.903 | 1915.833 | 0.337 | 74.3 MiB | 180516 |
| `sheetjs` | ran/lost vs Ascend median | 866.457 | 2503.080 | 0.481 | 201.8 MiB | 883673 |

Scoreboard result:

- Median scoreboard: sparse-wide group winner was `ascend-external-writer`.
- P95 scoreboard: sparse-wide group winner was `rust-xlsxwriter`.
- Both scoreboard commands exit nonzero for full-profile coverage because this
  is not a full `xlsx-write-sota` run. Missing/omitted full-profile libraries,
  unsupported rows, and blocked runners are not counted as wins.

Semantic comparability: all listed rows reopened successfully, matched the
expected one-sheet and 23,093-cell populated shape, and passed the harness'
external post-write semantic validation. Ascend, Excelize, SheetJS, and
XlsxWriter matched ordered semantic cell hashes. FastExcel Java and
rust_xlsxwriter passed semantic value validation but did not match ordered
semantic hashes, so their rows are useful lower-fidelity sparse value-write
comparisons, not byte/order-equivalent output claims.

Humble allowed wording:

> On the generated 5000 x 256 sparse-wide write row at commit `6595d42c`,
> Ascend's focused external repeat-15 run had the fastest median among the
> completed fastest comparable writers, but `rust-xlsxwriter` had the best p95,
> lower RSS, and smaller output. This is scoped sparse-write evidence and a
> tail-latency boundary, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend has the best sparse-wide write tail latency."
- "Ascend beats `rust-xlsxwriter` on p95 or memory for sparse-wide writes."
- "Ascend beats omitted, unsupported, blocked, or untested sparse-wide writers."
- "Ascend produces the smallest sparse-wide XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: do not optimize from the median win alone. Keep sparse-wide
write-speed wording scoped to median; if this workflow becomes a release p95
claim, profile Ascend's sparse-wide writer tail against `rust-xlsxwriter` before
making a production change.

## Cycle: Sparse Wide TS/JS/Rust Write Head-to-Head at `a328b573`

Classification: comparable external evidence plus stale-boundary update. This
refreshes the sparse-wide write row against the TS/JS and Rust floor runners.
Unlike the older `6595d42c` row, Ascend is now the median and p95 winner against
SheetJS, ExcelJS, and rust_xlsxwriter on the current clean commit. No production
optimization is justified from a winning row.

Workflow: generated XLSX write for sparse-wide values, 5000 rows x 256 columns,
23,093 populated cells.

Why it matters for release: sparse-wide workbooks model wide agent-generated
planning sheets and dashboards. The current benchmark floor policy prioritizes
TS/JS and Rust, and the historical sparse-wide claim boundary said
rust_xlsxwriter won p95. This run updates that boundary for current evidence
without broadening the claim beyond this focused row.

Public/tracked-clean input: `competitive-io` generated the `sparse-wide`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `a328b573`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-sparse-js-rust-current-a328b573 a328b573
cd /private/tmp/ascend-write-sparse-js-rust-current-a328b573
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,rust-xlsxwriter --workload sparse-wide --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15.json 2> /private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15.json --json --metric p95Ms --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15-p95-scoreboard.json
```

Environment:

- Commit: `a328b573`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-sparse-js-rust-current-a328b573`; `git status
  --short --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload sparse-wide`, `validationMode each`,
  `repeat 15`, `warmup 3`.
- Runner versions: SheetJS `0.18.5`, ExcelJS `4.4.0`, rust_xlsxwriter `0.1.0`.

Raw output:

```text
/private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15.json
/private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15-time.txt
/private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15-scoreboard.json
/private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15-p95-scoreboard.json
```

Focused TS/JS/Rust row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won median and p95 | 15.530 | 20.085 | 0.134 | 169.9 MiB | 228209 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 21.016 | 23.789 | 0.048 | 52.7 MiB | 175581 |
| `sheetjs` | ran/lost vs Ascend | 402.837 | 593.206 | 0.185 | 351.3 MiB | 883673 |
| `exceljs` | ran/lost vs Ascend | 3892.681 | 7766.345 | 0.353 | 1563.6 MiB | 184376 |

Process-level `/usr/bin/time -l` for the full command reported `103.78 real`,
2,629,320,704 bytes maximum resident set size, and 118,916,008 bytes peak
memory footprint.

Scoreboard result:

- Median scoreboard: sparse-wide group winner was `ascend-external-writer`;
  `leaderFailures: []` and `profileLeaderFailures: []`.
- P95 scoreboard: sparse-wide group winner was `ascend-external-writer`;
  `leaderFailures: []` and `profileLeaderFailures: []`.
- Both scoreboard commands exit nonzero for full-profile coverage because this
  is not a full `xlsx-write-sota` run. Each reports 59 coverage failures for
  missing workloads and omitted writers. Those rows are not wins.

Semantic comparability: all four rows reopened successfully, matched the
expected one-sheet and 23,093 populated-cell shape, and passed semantic cell
value validation. Ascend and SheetJS matched ordered semantic cell hashes.
ExcelJS and rust_xlsxwriter matched sorted semantic values but not ordered
semantic hashes, so their rows are useful value-write timing evidence but not
byte/order-equivalent output evidence. Ascend wins median and p95 here, but
rust_xlsxwriter and ExcelJS emit smaller XLSX files and rust_xlsxwriter uses
less RSS.

Humble allowed wording:

> On the generated 5000 x 256 `sparse-wide` write row at commit `a328b573`,
> Ascend's focused external repeat-15 run was faster by median and p95 than
> SheetJS `0.18.5`, ExcelJS `4.4.0`, and rust_xlsxwriter `0.1.0`, with all rows
> passing semantic value validation. This updates the older sparse-wide
> rust_xlsxwriter p95 boundary for current evidence, but remains scoped
> TS/JS/Rust generated-write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every TS/JS or Rust writer on every workload."
- "Ascend proves byte/order-equivalent output against ExcelJS or rust_xlsxwriter."
- "Ascend produces the smallest sparse-wide XLSX."
- "Ascend uses less memory than rust_xlsxwriter on sparse-wide writes."

Next action: keep TS/JS/Rust benchmark work focused on rows where Ascend lacks
current evidence or loses. Do not re-optimize this sparse-wide row unless a
current JS/Rust rerun regresses its median or p95.

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

## Cycle: Plain Text TS/JS/Rust Write Head-to-Head at `1bd995e5`

Classification: comparable external evidence plus defer. This refreshes the
plain-text generated write row against the TS/JS and Rust floor runners. Ascend
is the median and p95 winner against SheetJS, ExcelJS, and rust_xlsxwriter in
this row. No production optimization is justified from a winning row.

Workflow: generated XLSX write for plain text values, 2000 rows x 20 columns.

Why it matters for release: plain text generated exports are a basic
agent-produced workbook shape, and the current benchmark floor policy
prioritizes TS/JS and Rust comparisons. This run provides a clean current
repeat-15 row for SheetJS, ExcelJS, and rust_xlsxwriter together after the
writer metadata-key optimization and recent claim-matrix commits.

Public/tracked-clean input: `competitive-io` generated the `plain-text`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `1bd995e5`. No private corpus or local research
workbook was used. The generated sheet has one sheet, 2,000 rows, 20 columns,
and 40,000 logical cells.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-plain-js-rust-current-1bd995e5 1bd995e5
cd /private/tmp/ascend-write-plain-js-rust-current-1bd995e5
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,rust-xlsxwriter --workload plain-text --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15.json 2> /private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15.json --json --metric p95Ms --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15-p95-scoreboard.json
```

Environment:

- Commit: `1bd995e5`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-plain-js-rust-current-1bd995e5`; `git status
  --short --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload plain-text`, `validationMode each`,
  `repeat 15`, `warmup 3`.
- Runner versions: SheetJS `0.18.5`, ExcelJS `4.4.0`, rust_xlsxwriter `0.1.0`.

Raw output:

```text
/private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15.json
/private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15-time.txt
/private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15-scoreboard.json
/private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15-p95-scoreboard.json
```

Focused TS/JS/Rust row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won median and p95 | 3.898 | 4.080 | 0.033 | 96.9 MiB | 169097 |
| `sheetjs` | ran/lost vs Ascend | 29.548 | 42.209 | 0.124 | 278.4 MiB | 1832541 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 29.865 | 33.282 | 0.038 | 28.7 MiB | 229139 |
| `exceljs` | ran/lost vs Ascend | 91.269 | 112.351 | 0.067 | 302.8 MiB | 232106 |

Process-level `/usr/bin/time -l` for the full command reported `16.31 real`,
341,327,872 bytes maximum resident set size, and 129,762,144 bytes peak memory
footprint.

Scoreboard result:

- Median scoreboard: plain-text group winner was `ascend-external-writer`;
  `leaderFailures: []` and `profileLeaderFailures: []`.
- P95 scoreboard: plain-text group winner was `ascend-external-writer`;
  `leaderFailures: []` and `profileLeaderFailures: []`.
- Both scoreboard commands exit nonzero for full-profile coverage because this
  is not a full `xlsx-write-sota` run. Each reports 59 coverage failures for
  missing workloads and omitted writers. Those rows are not wins.

Semantic comparability: all four rows reopened successfully, matched one sheet
and 40,000 cells, and passed semantic cell value validation. Ascend and SheetJS
matched ordered semantic cell hashes. ExcelJS and rust_xlsxwriter matched
sorted semantic values but not ordered semantic hashes, so their rows are useful
value-write timing evidence but not byte/order-equivalent output evidence.
Ascend wins median and p95 here, but rust_xlsxwriter uses less RSS.

Humble allowed wording:

> On the generated 2000 x 20 `plain-text` write row at commit `1bd995e5`,
> Ascend's focused external repeat-15 run was faster by median and p95 than
> SheetJS `0.18.5`, ExcelJS `4.4.0`, and rust_xlsxwriter `0.1.0`, with all rows
> passing semantic value validation. This is scoped TS/JS/Rust generated-write
> evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every TS/JS or Rust writer on every workload."
- "Ascend proves byte/order-equivalent output against ExcelJS or rust_xlsxwriter."
- "Ascend produces the smallest plain-text XLSX."
- "Ascend uses less memory than rust_xlsxwriter on plain-text writes."

Next action: keep TS/JS/Rust benchmark work focused on rows where Ascend lacks
current evidence or loses. Do not re-optimize this plain-text row unless a
current JS/Rust rerun regresses its median or p95.

## Cycle: Plain Text Write Current Split Row at `91dabea8`

Classification: comparable external evidence plus blocked broad-row boundary.
The all-runner current plain-text generated-write row was killed as an opaque
zero-byte blocker, then a focused fastest/profile-critical split row completed.
Ascend is the median and p95 winner among the completed focused comparable
writers, but the Ascend, FastExcel Java, and SheetJS rows are noisy. This is
scoped plain-text write evidence, not a full write-values or `xlsx-write-sota`
claim.

Workflow: generated XLSX write for plain text values, 2000 rows x 20 columns.

Why it matters for release: text-heavy generated exports are a common
agent-produced workbook output, and `plain-text` is one of the value-write rows
required by the existing `xlsx-write-sota` profile.

Public/tracked-clean input: `competitive-io` generated the `plain-text`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `91dabea8`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-plain-current-91dabea8 91dabea8fc84
cd /private/tmp/ascend-write-plain-current-91dabea8
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-plain-current-91dabea8-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,xlsxwriter,xlsxwriter-constant-memory,pyexcelerate,pyexcelerate-range,pyexcelerate-cell,openpyxl,openpyxl-write-only,apache-poi,closedxml,rust-xlsxwriter,excelize,fastexcel-java --workload plain-text --repeat 5 --warmup 1 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-plain-current-91dabea8-runs/write-plain-text-repeat5.json
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,excelize,rust-xlsxwriter,sheetjs,fastexcel-java,xlsxwriter,xlsxwriter-constant-memory --workload plain-text --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-plain-current-91dabea8-runs/write-plain-text-fastest-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-plain-current-91dabea8-runs/write-plain-text-fastest-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-plain-current-91dabea8-runs/write-plain-text-fastest-repeat15-scoreboard.json
```

Environment:

- Commit: `91dabea8fc84eac74efca3efd0081a0e44766748`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-plain-current-91dabea8`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload plain-text`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-plain-current-91dabea8-runs/write-plain-text-repeat5.json
/private/tmp/ascend-write-plain-current-91dabea8-runs/write-plain-text-fastest-repeat15.json
/private/tmp/ascend-write-plain-current-91dabea8-runs/write-plain-text-fastest-repeat15-scoreboard.json
```

Blocked broad row:

| Artifact | Status | Sample count | Median | P95 | CV/noise | Memory/size |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `write-plain-text-repeat5.json` | blocked, killed before JSON emission | 0 | n/a | n/a | n/a | n/a |

Focused fastest/profile-critical split row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 12.238 | 60.870 | 0.805 | 98.4 MiB | 169097 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 93.945 | 229.973 | 0.397 | 31.2 MiB | 229139 |
| `fastexcel-java` | ran/lost vs Ascend | 123.541 | 568.597 | 0.816 | 1456.0 MiB | 227254 |
| `excelize` | ran/lost vs Ascend | 138.027 | 235.307 | 0.474 | 23.2 MiB | 142890 |
| `sheetjs` | ran/lost vs Ascend | 151.009 | 449.305 | 0.621 | 194.4 MiB | 1832541 |
| `xlsxwriter` | ran/lost vs Ascend | 444.135 | 1062.576 | 0.405 | 94.4 MiB | 230175 |
| `xlsxwriter-constant-memory` | ran/lost vs Ascend | 467.383 | 812.873 | 0.264 | 56.4 MiB | 141062 |

Scoreboard result:

- Focused repeat-15 plain-text split row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The scoreboard command exits nonzero for full-profile coverage because this is
  a single split row. Omitted writers from the blocked broad row are not counted
  as Ascend wins.

Semantic comparability: all listed focused rows reopened successfully, matched
the expected sheet and cell-count shape, and passed the benchmark's write
correctness gate. The semantic hash assertions differ by writer ordering lane,
so this row supports value-write comparability, not byte/order-equivalent output
claims. Ascend is not the smallest output and uses more RSS than Excelize,
rust_xlsxwriter, and XlsxWriter constant-memory.

Humble allowed wording:

> On the generated plain-text write row at commit `91dabea8`, Ascend's focused
> external repeat-15 split run had the fastest median and p95 among the completed
> focused comparable writers. The broader all-runner plain-text row was blocked
> before JSON emission, so this is scoped plain-text write evidence, not a broad
> `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every plain-text writer."
- "Ascend has a low-noise plain-text tail win."
- "Ascend beats omitted, unsupported, blocked, or untested plain-text writers."
- "Ascend produces the smallest plain-text XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: defer production optimization for plain-text from this focused
median/p95 win, keep the broad plain-text row blocked, and continue splitting
the write-values profile into attributable rows.

## Cycle: Plain Text ClosedXML/NPOI Write Coverage at `e0c41fe5`

Classification: comparable external evidence plus blocker update. ClosedXML and
NPOI now run and pass the generated `plain-text` value-write validation in the
current clean worktree. This updates the older ClosedXML
`CSSM_ModuleLoad()`/unavailable boundary for this focused row, but it is still
not broad `xlsx-write-sota` coverage.

Workflow: generated XLSX write for plain text values, 2000 rows x 20 columns,
focused on the previously blocked .NET writer coverage.

Why it matters for release: broad write speed wording has been blocked partly by
opaque runner failures and missing external writer rows. This run converts the
current `plain-text` ClosedXML and NPOI status from blocked/unknown to measured,
validated, and slower than Ascend. It does not resolve other value-write
workloads or feature-rich/profile coverage.

Public/tracked-clean input: `competitive-io` generated the `plain-text`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `e0c41fe5`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-closedxml-current-e0c41fe5 e0c41fe5
cd /private/tmp/ascend-write-closedxml-current-e0c41fe5
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-closedxml-current-e0c41fe5-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,closedxml,npoi --workload plain-text --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-closedxml-current-e0c41fe5-runs/write-plain-text-closedxml-npoi-repeat15.json 2> /private/tmp/ascend-write-closedxml-current-e0c41fe5-runs/write-plain-text-closedxml-npoi-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-closedxml-current-e0c41fe5-runs/write-plain-text-closedxml-npoi-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-closedxml-current-e0c41fe5-runs/write-plain-text-closedxml-npoi-repeat15-scoreboard.json
```

Environment:

- Commit: `e0c41fe58aa90f3b40c62d5bf396caea5cd7db7b`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-closedxml-current-e0c41fe5`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload plain-text`, `validationMode each`,
  `repeat 15`, `warmup 3`, `ACCEPT_NPOI_OSMF_LICENSE=1`.

Raw output:

```text
/private/tmp/ascend-write-closedxml-current-e0c41fe5-runs/write-plain-text-closedxml-npoi-repeat15.json
/private/tmp/ascend-write-closedxml-current-e0c41fe5-runs/write-plain-text-closedxml-npoi-repeat15-time.txt
/private/tmp/ascend-write-closedxml-current-e0c41fe5-runs/write-plain-text-closedxml-npoi-repeat15-scoreboard.json
```

Focused .NET writer coverage row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes | Runtime |
| --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| `ascend-external-writer` | ran/won | 4.008 | 4.203 | 0.029 | 98.3 MiB | 169097 | Bun `1.3.13`, workspace runner |
| `closedxml` | ran/lost vs Ascend | 102.203 | 189.116 | 0.323 | 167.5 MiB | 224794 | .NET, ClosedXML `0.105.0.0` |
| `npoi` | ran/lost vs Ascend | 201.248 | 312.078 | 0.244 | 150.4 MiB | 224417 | .NET, NPOI `2.8.0.0` |

Scoreboard result:

- Focused `plain-text` group winner: `ascend-external-writer`.
- `leaderFailures: []` and `profileLeaderFailures: []`.
- Full `xlsx-write-sota` coverage still fails because this row omits other
  workloads and required writers.

Semantic comparability: all three rows wrote the same generated plain-text
sheet, reopened successfully, matched the benchmark's shape and value
assertions, and reported `correctnessStatus: pass` with `rankingEligible: true`.
This row proves current ClosedXML and NPOI availability for this generated
plain-text value-write workflow only. It does not prove coverage for
`dense-values`, `sparse-wide`, `string-heavy`, styles, formulas, tables,
feature-rich metadata, file-size leadership, or broad write SOTA.

Humble allowed wording:

> On the generated 2000 x 20 `plain-text` write row at current commit
> `e0c41fe5`, ClosedXML and NPOI both ran and passed validation, and Ascend was
> faster by median and p95 on this focused row. This updates the older ClosedXML
> runner-blocked boundary for `plain-text`, but it is not broad
> `xlsx-write-sota` coverage.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats ClosedXML or NPOI on every write workload."
- "ClosedXML and NPOI broad write coverage is complete."
- "Ascend produces the smallest plain-text XLSX."
- Any wording that treats this focused `plain-text` row as full write-profile
  promotion.

Next action: use current ClosedXML/NPOI runner availability to split additional
write-values blockers beyond `plain-text` and `string-heavy` only where they
close named coverage gaps; keep broad write wording downgraded until
multi-workload/full-profile coverage is complete.

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

## Cycle: String Heavy TS/JS/Rust Write Head-to-Head at `05de2c46`

Classification: comparable external evidence plus stale-boundary update. This
refreshes the generated string-heavy write row against the primary TS/JS and
Rust floor after the string-heavy optimization and later writer/calc commits.
Ascend is the median and p95 winner against SheetJS, ExcelJS, and
rust_xlsxwriter in this row, so no production optimization is justified from
the current JS/Rust floor.

Workflow: generated XLSX write for varied string values, 2000 rows x 20 columns,
40,000 populated cells.

Why it matters for release: the older string-heavy matrix row had a noisy
repeat-15 tail and omitted ExcelJS from the focused fastest-writer rerun. The
user explicitly prioritizes heads-up JS/TS comparisons and Rust as the minimum
performance floor, so this run closes that exact evidence gap for current
wording.

Public/tracked-clean input: `competitive-io` generated the `string-heavy`
`source-mode generated-write` workload from tracked benchmark code at commit
`05de2c46`. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-string-js-rust-current-05de2c46 05de2c463a455c57d6d85284e4808cbfd87b9c5a
cd /private/tmp/ascend-write-string-js-rust-current-05de2c46
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-string-js-rust-current-05de2c46-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,rust-xlsxwriter --workload string-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15.json 2> /private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15.json --json --metric p95Ms --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15-p95-scoreboard.json
```

Environment:

- Commit: `05de2c463a455c57d6d85284e4808cbfd87b9c5a`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-string-js-rust-current-05de2c46`; `git status
  --short --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `22.22.0`
- Rust: `rustc 1.91.1`
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runner versions: SheetJS `0.18.5`, ExcelJS `4.4.0`, rust_xlsxwriter runner
  `0.1.0` using `rust_xlsxwriter` `0.94.0`.
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload string-heavy`, `repeat 15`,
  `warmup 3`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15.json
/private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15-time.txt
/private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15-scoreboard.json
/private/tmp/ascend-write-string-js-rust-current-05de2c46-runs/write-string-heavy-js-rust-repeat15-p95-scoreboard.json
```

Focused TS/JS/Rust writer row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won median and p95 | 4.056 | 4.198 | 0.028 | 90.8 MiB | 201984 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 27.591 | 28.267 | 0.022 | 23.7 MiB | 237837 |
| `sheetjs` | ran/lost vs Ascend | 33.028 | 34.293 | 0.024 | 272.2 MiB | 2016032 |
| `exceljs` | ran/lost vs Ascend | 83.537 | 114.802 | 0.108 | 282.1 MiB | 240319 |

Process-level `/usr/bin/time -l`: `16.00 real`, `6.01 user`, `0.55 sys`,
`318128128` maximum resident set size, `123110264` peak memory footprint.

Scoreboard result:

- Median scoreboard: string-heavy group winner was `ascend-external-writer`;
  `leaderFailures: []`, `profileLeaderFailures: []`.
- P95 scoreboard: string-heavy group winner was `ascend-external-writer`;
  `leaderFailures: []`, `profileLeaderFailures: []`.
- Full `xlsx-write-sota` coverage still fails, with 59 coverage failures,
  because this is a focused row rather than a full-profile promotion run.

Semantic comparability: all four rows write the same generated string-heavy
sheet and pass external post-write semantic validation for one sheet and 40,000
cells. Ascend and the JS writers match ordered semantic cell hashes;
rust_xlsxwriter matches sorted semantic values but not ordered cell hashes.
Memory and size tradeoffs remain material: rust_xlsxwriter uses less RSS, and
ExcelJS/rust_xlsxwriter emit larger than Ascend but much smaller than SheetJS.

Humble allowed wording:

> On the generated 2000 x 20 string-heavy write row at commit `05de2c46`,
> Ascend's external generated writer was faster by median and p95 than SheetJS
> `0.18.5`, ExcelJS `4.4.0`, and rust_xlsxwriter `0.94.0`, with all rows
> passing post-write semantic validation. This supersedes the older noisy
> string-heavy p95 boundary for current TS/JS/Rust wording, but remains scoped
> generated value-write evidence.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every TS/JS or Rust writer on every workload."
- "Ascend uses less memory than rust_xlsxwriter on string-heavy writes."
- "Ascend produces the smallest string-heavy XLSX."
- "Ascend proves byte/order-equivalent output against rust_xlsxwriter."

Next action: defer production optimization from this winning row. Continue with
another blocked JS/Rust floor row, a true current loss, or a full-profile write
gate split that can produce complete attributable coverage.

## Cycle: Plain Text Workbook Writer Metadata-Key Optimization at `fd616906`

Classification: validated optimization. The workbook-buffered sheet writer was
computing a per-cell formula metadata key even when the sheet had no stored
formula text and no preserved formula cell metadata. The patched writer checks
those maps once per sheet and only computes the key when a lookup is possible.

Workflow: generated XLSX write for plain text values, 2000 rows x 20 columns,
using the workbook-buffered writer path.

Why it matters for release: this is the small workbook writer path used by
normal `writeXlsx` value exports. The result improves a user-visible generated
write workflow without changing workbook semantics. It is not a broad
`xlsx-write-sota` claim because it is a focused phase A/B and the full write
profile still has coverage blockers.

Public/tracked-clean input: `xlsx-write-phase` generated the `plain-text`
workload from tracked benchmark code in a clean detached worktree at base commit
`fd6169060eba7e45edbdf83d1f18d97ddcda11b4`. The patched run applied only the
production diff for `packages/io-xlsx/src/writer/sheet.ts`. No private corpus or
local research workbook was used.

Production change:

- `packages/io-xlsx/src/writer/sheet.ts` now computes `formulaStorageKey(row,
  col)` only when `storedFormulaText` or `preservedCellMetadata` is non-empty.
- Existing writer tests continue to prove stored formula text and formula cell
  metadata survive unrelated dirty-sheet edits.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-writer-key-current-fd616906 fd6169060eba7e45edbdf83d1f18d97ddcda11b4
cd /private/tmp/ascend-writer-key-current-fd616906
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-writer-key-current-fd616906-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/xlsx-write-phase.ts --workload plain-text --rows 2000 --cols 20 --repeat 40 --warmup 5 --gc-between-samples --json > /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-baseline-repeat40.json 2> /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-baseline-repeat40-time.txt
git apply /private/tmp/ascend-writer-sheet-metadata-key.patch
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/xlsx-write-phase.ts --workload plain-text --rows 2000 --cols 20 --repeat 40 --warmup 5 --gc-between-samples --json > /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-repeat40.json 2> /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-repeat40-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/xlsx-write-phase.ts --workload plain-text --rows 2000 --cols 20 --repeat 3 --warmup 1 --gc-between-samples --validate --json > /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-validate.json 2> /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-validate-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun test packages/io-xlsx/src/writer/writer.test.ts -t 'preserves original stored formula text when unrelated edits dirty the sheet|preserves formula cell metadata when unrelated edits dirty the sheet'
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,rust-xlsxwriter,excelize,sheetjs,fastexcel-java --workload plain-text --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-external-repeat15.json 2> /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-external-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-external-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-external-repeat15-scoreboard.json
```

Environment:

- Base commit: `fd6169060eba7e45edbdf83d1f18d97ddcda11b4`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-writer-key-current-fd616906`; patched worktree reported
  only `packages/io-xlsx/src/writer/sheet.ts` dirty.
- Bun runtime: `1.3.13`
- Node CLI: `v26.0.0`; competitive harness runtime metadata reported Node
  `24.3.0`
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runtime profile: `workload plain-text`, `rows 2000`, `cols 20`, `repeat 40`,
  `warmup 5`, `gcBetweenSamples true`, `streaming false`,
  `writerPath workbook-buffered`.

Raw output:

```text
/private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-baseline-repeat40.json
/private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-baseline-repeat40-time.txt
/private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-repeat40.json
/private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-repeat40-time.txt
/private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-validate.json
/private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-external-repeat15.json
/private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-external-repeat15-scoreboard.json
```

Focused workbook-buffered phase A/B, repeat 40 after 5 warmups:

| Run | Write median ms | Write p95 ms | Write CV | Total median ms | Total p95 ms | Total CV | Median throughput | Peak RSS | Bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Baseline | 4.780 | 6.607 | 0.211 | 8.126 | 10.175 | 0.135 | 8.367 M cells/s | 161.8 MiB | 176828 |
| Patched | 4.094 | 4.762 | 0.076 | 7.024 | 7.999 | 0.058 | 9.771 M cells/s | 161.1 MiB | 176828 |

Comparison: patched write median improved by `14.36%`, write p95 improved by
`27.92%`, and total median improved by `13.57%`. Output size was unchanged.
Process-level `/usr/bin/time -l` reported 170,082,304 bytes maximum resident set
size for baseline and 169,394,176 bytes for patched.

Correctness: patched `--validate` repeat 3 passed reopen validation for the same
40,000 cells, with `validateMedianMs 68.985` and unchanged `176828` median
bytes. Focused writer tests passed:

```text
packages/io-xlsx/src/writer/writer.test.ts:
(pass) writeXlsx > preserves original stored formula text when unrelated edits dirty the sheet
(pass) writeXlsx > preserves formula cell metadata when unrelated edits dirty the sheet
```

Current external context, patched repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 4.319 | 4.615 | 0.033 | 97.5 MiB | 169097 |
| `excelize` | ran/lost vs Ascend | 26.881 | 52.524 | 0.273 | 23.5 MiB | 142890 |
| `sheetjs` | ran/lost vs Ascend | 28.253 | 30.247 | 0.041 | 280.7 MiB | 1832541 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 33.953 | 36.438 | 0.034 | 26.7 MiB | 229138 |
| `fastexcel-java` | ran/lost vs Ascend | 46.587 | 239.669 | 0.851 | 1584.0 MiB | 227254 |

Scoreboard result:

- Focused patched external row: `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The full `xlsx-write-sota` gate still fails coverage because this row omits
  other workloads and several required writers for `plain-text`.

Semantic comparability: the phase A/B measures the same generated workbook
writer path and validates reopen semantics after the patch. The external row
uses generated-write post-operation validation for the same 40,000-cell plain
text shape, but it is current patched context, not a before/after external A/B.
Memory and file size are not normalized across external writers; Excelize and
rust_xlsxwriter use less RSS, while Ascend is faster by median and p95 on the
completed focused row.

Humble allowed wording:

> On the generated 2000 x 20 plain-text workbook-buffered write phase at base
> commit `fd616906`, skipping impossible formula metadata key lookups improved
> Ascend's median write time by 14.36% and p95 by 27.92% without changing output
> size or reopen validation. A current patched external plain-text row remained
> faster by median and p95 than the completed Excelize, SheetJS,
> rust_xlsxwriter, and FastExcel Java rows.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every generated XLSX writer."
- "Ascend improved every write workload."
- "Ascend has the smallest plain-text XLSX."
- Any wording that treats this focused phase A/B as a full-profile promotion.

Next action: do not optimize this same plain-text workbook-buffered row again
unless a current full write gate regresses it. Continue with remaining broad
write coverage blockers, feature-rich comparability, or real-workbook workflow
latency.

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

## Cycle: String Heavy ClosedXML/NPOI Write Coverage at `6cc5076f`

Classification: comparable external evidence plus blocker update. ClosedXML
and NPOI now run and pass validation on the generated `string-heavy` write row,
so the earlier omitted/unavailable-runner wording is historical for this focused
row. Ascend is the median and p95 winner against both .NET writers here. This
is still not a broad `xlsx-write-sota` claim.

Workflow: generated XLSX write for varied string values, 2000 rows x 20 columns.

Why it matters for release: `string-heavy` is a required value-write workflow in
the existing `xlsx-write-sota` profile, and the previous current fastest row
omitted ClosedXML and NPOI. This run closes that specific .NET writer coverage
gap for the generated string-write shape without promoting the full write
profile.

Public/tracked-clean input: `competitive-io` generated the `string-heavy`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `6cc5076f`. No private corpus or local research
workbook was used. The generated sheet has one sheet, 2,000 rows, 20 columns,
and 40,000 logical cells.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-dotnet-current-6cc5076f 6cc5076f85ce995a55bb6d83a5c8d31a9e7540b1
cd /private/tmp/ascend-write-dotnet-current-6cc5076f
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-dotnet-current-6cc5076f-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,closedxml,npoi --workload string-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-dotnet-current-6cc5076f-runs/write-string-heavy-closedxml-npoi-repeat15.json 2> /private/tmp/ascend-write-dotnet-current-6cc5076f-runs/write-string-heavy-closedxml-npoi-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-dotnet-current-6cc5076f-runs/write-string-heavy-closedxml-npoi-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-dotnet-current-6cc5076f-runs/write-string-heavy-closedxml-npoi-repeat15-scoreboard.json
```

Environment:

- Commit: `6cc5076f85ce995a55bb6d83a5c8d31a9e7540b1`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-dotnet-current-6cc5076f`; `git status --short
  --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload string-heavy`, `validationMode each`,
  `repeat 15`, `warmup 3`.
- .NET runners: ClosedXML `0.105.0.0`; NPOI `2.8.0.0`, with
  `ACCEPT_NPOI_OSMF_LICENSE=1`.

Raw output:

```text
/private/tmp/ascend-write-dotnet-current-6cc5076f-runs/write-string-heavy-closedxml-npoi-repeat15.json
/private/tmp/ascend-write-dotnet-current-6cc5076f-runs/write-string-heavy-closedxml-npoi-repeat15-time.txt
/private/tmp/ascend-write-dotnet-current-6cc5076f-runs/write-string-heavy-closedxml-npoi-repeat15-scoreboard.json
```

Focused .NET writer row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 4.131 | 4.820 | 0.064 | 91.9 MiB | 201984 |
| `closedxml` | ran/lost vs Ascend | 113.223 | 317.533 | 0.502 | 140.0 MiB | 246941 |
| `npoi` | ran/lost vs Ascend | 206.278 | 480.695 | 0.376 | 146.6 MiB | 234032 |

Process-level `/usr/bin/time -l` for the full command reported `26.62 real`,
209,207,296 bytes maximum resident set size, and 120,062,840 bytes peak memory
footprint.

Scoreboard result:

- Focused repeat-15 .NET writer row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- Full `xlsx-write-sota` coverage still fails, with 61 coverage failures,
  because this row intentionally covers only `string-heavy` and only Ascend,
  ClosedXML, and NPOI. Missing workloads and omitted writers are not wins.

Semantic comparability: all three rows reopened successfully, matched one sheet
and 40,000 cells, and passed semantic cell value validation. Ascend matched
ordered semantic cell hashes. ClosedXML and NPOI matched sorted semantic values
but not ordered semantic hashes, so the .NET rows are comparable value-write
timing evidence, not byte/order-equivalent output claims. Ascend wins median
and p95 here but does not produce the smallest XLSX.

Humble allowed wording:

> On the generated 2000 x 20 `string-heavy` write row at commit `6cc5076f`,
> Ascend's focused external repeat-15 run was faster by median and p95 than
> ClosedXML `0.105.0.0` and NPOI `2.8.0.0`, with all three rows passing
> semantic value validation. This is scoped generated string-write evidence, not
> a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats ClosedXML or NPOI on every write workload."
- "Ascend proves byte/order-equivalent output against ClosedXML or NPOI."
- "Ascend produces the smallest string-heavy XLSX."
- "ClosedXML and NPOI broad write coverage is complete."

Next action: use current ClosedXML/NPOI runner availability to split any
remaining write-values blockers only where they close named coverage gaps. Do
not revisit `string-heavy` optimization unless a current full-profile or
fastest-writer row regresses it.

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

## Cycle: Styles Heavy TS/JS/Rust Write Head-to-Head at `38feccee`

Classification: comparable external evidence plus defer. This refreshes the
styles-heavy generated write row against the TS/JS and Rust floor runners.
Ascend is the median and p95 winner against SheetJS, ExcelJS, and
rust_xlsxwriter in this row. No production optimization is justified from a
winning row.

Workflow: generated XLSX write for styled numeric cells, 2000 rows x 20 columns.

Why it matters for release: styled generated workbook export is a user-visible
agent commit/export workflow, and the benchmark floor policy prioritizes TS/JS
and Rust comparisons. This run adds SheetJS and ExcelJS to the focused
styles-heavy floor evidence instead of relying only on older mixed fastest-row
evidence.

Public/tracked-clean input: `competitive-io` generated the `styles-heavy`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `38feccee`. No private corpus or local research
workbook was used. The generated sheet has one sheet, 2,000 rows, 20 columns,
and 40,000 logical cells.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-styles-js-rust-current-38feccee 38feccee
cd /private/tmp/ascend-write-styles-js-rust-current-38feccee
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-styles-js-rust-current-38feccee-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,rust-xlsxwriter --workload styles-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15.json 2> /private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15.json --json --metric p95Ms --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15-p95-scoreboard.json
```

Environment:

- Commit: `38feccee`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-styles-js-rust-current-38feccee`; `git status
  --short --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload styles-heavy`, `validationMode each`,
  `repeat 15`, `warmup 3`.
- Runner versions: SheetJS `0.18.5`, ExcelJS `4.4.0`, rust_xlsxwriter `0.1.0`.

Raw output:

```text
/private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15.json
/private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15-time.txt
/private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15-scoreboard.json
/private/tmp/ascend-write-styles-js-rust-current-38feccee-runs/write-styles-heavy-js-rust-repeat15-p95-scoreboard.json
```

Focused TS/JS/Rust row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won median and p95 | 9.450 | 12.501 | 0.095 | 177.8 MiB | 434417 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 21.331 | 22.837 | 0.026 | 18.7 MiB | 227731 |
| `sheetjs` | ran/lost vs Ascend | 28.374 | 32.362 | 0.062 | 245.3 MiB | 1165452 |
| `exceljs` | ran/lost vs Ascend | 64.839 | 130.125 | 0.248 | 248.5 MiB | 215294 |

Process-level `/usr/bin/time -l` for the full command reported `15.61 real`,
277,938,176 bytes maximum resident set size, and 126,059,360 bytes peak memory
footprint.

Scoreboard result:

- Median scoreboard: styles-heavy group winner was `ascend-external-writer`;
  `leaderFailures: []` and `profileLeaderFailures: []`.
- P95 scoreboard: styles-heavy group winner was `ascend-external-writer`;
  `leaderFailures: []` and `profileLeaderFailures: []`.
- Both scoreboard commands exit nonzero for full-profile coverage because this
  is not a full `xlsx-write-sota` run. Each reports 61 coverage failures for
  missing workloads and omitted writers. Those rows are not wins.

Semantic comparability: all four rows reopened successfully, matched one sheet
and 40,000 cells, and passed semantic cell value validation. Ascend, SheetJS,
and ExcelJS matched ordered semantic cell hashes. rust_xlsxwriter matched
sorted semantic values but not ordered semantic hashes, so its row is useful
value-write timing evidence but not byte/order-equivalent output evidence. This
row validates value semantics, not style-fidelity equivalence across writer
libraries. Ascend wins median and p95 here, but rust_xlsxwriter and ExcelJS emit
smaller XLSX files and rust_xlsxwriter uses less RSS.

Humble allowed wording:

> On the generated 2000 x 20 `styles-heavy` write row at commit `38feccee`,
> Ascend's focused external repeat-15 run was faster by median and p95 than
> SheetJS `0.18.5`, ExcelJS `4.4.0`, and rust_xlsxwriter `0.1.0`, with all rows
> passing semantic value validation. This is scoped TS/JS/Rust generated-write
> evidence, not a broad `xlsx-write-sota` or cross-library style-fidelity claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every TS/JS or Rust writer on every workload."
- "Ascend proves style-fidelity equivalence against SheetJS, ExcelJS, or rust_xlsxwriter."
- "Ascend produces the smallest styles-heavy XLSX."
- "Ascend uses less memory than rust_xlsxwriter on styles-heavy writes."

Next action: keep TS/JS/Rust benchmark work focused on rows where Ascend lacks
current evidence or loses. Do not re-optimize this styles-heavy row unless a
current JS/Rust rerun regresses its median or p95.

## Cycle: Formula Heavy Write Fastest Comparable Row at `7118f208`

Classification: comparable external evidence. Ascend is the median and p95
winner among the completed formula-capable generated XLSX writers in this row.
The full `xlsx-write-sota` profile is still not promotable from this single
workload, and no production optimization is justified from a winning row.

Workflow: generated XLSX write for formula-heavy workbooks, 2000 rows x 20
columns.

Why it matters for release: formula-bearing workbook export is a visible
commit/export workflow for agent-authored sheets. This row checks whether
Ascend's write path remains competitive when most cells are formulas rather
than plain values or styles.

Public/tracked-clean input: `competitive-io` generated the `formula-heavy`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `7118f208`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-formula-current-7118f208 7118f2086545
cd /private/tmp/ascend-write-formula-current-7118f208
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-formula-current-7118f208-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,xlsxwriter,xlsxwriter-constant-memory,rust-xlsxwriter --workload formula-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-formula-current-7118f208-runs/write-formula-heavy-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-formula-current-7118f208-runs/write-formula-heavy-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-formula-current-7118f208-runs/write-formula-heavy-repeat15-scoreboard.json
```

Environment:

- Commit: `7118f20865456199b7e37f848b9df6018f6dee2d`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-formula-current-7118f208`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload formula-heavy`,
  `validationMode each`, `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-write-formula-current-7118f208-runs/write-formula-heavy-repeat15.json
/private/tmp/ascend-write-formula-current-7118f208-runs/write-formula-heavy-repeat15-scoreboard.json
```

Focused formula-capable writer rerun, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 41.758 | 85.659 | 0.293 | 196.9 MiB | 405795 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 72.837 | 100.303 | 0.163 | 23.6 MiB | 245490 |
| `xlsxwriter` | ran/lost vs Ascend | 1238.221 | 2176.396 | 0.243 | 117.0 MiB | 234257 |
| `xlsxwriter-constant-memory` | ran/lost vs Ascend | 1351.949 | 1802.284 | 0.144 | 96.4 MiB | 234041 |

Scoreboard result:

- Focused repeat-15 formula-capable row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The scoreboard command exits nonzero for full-profile coverage because this is
  not a full `xlsx-write-sota` run. Coverage failures and omitted/unsupported
  formula writers are not counted as Ascend wins.

Semantic comparability: all listed rows reopened successfully, matched the
expected one-sheet and 40,000-cell shape, matched the expected 36,000 formula
cells, and passed sorted semantic cell value validation. XlsxWriter,
XlsxWriter constant-memory, and rust_xlsxwriter did not match ordered semantic
cell value hashes, so their rows are useful formula-write comparisons but not
byte/order-equivalent output claims. Ascend uses more RSS and emits a larger
XLSX than the compared formula writers.

Humble allowed wording:

> On the generated 2000 x 20 formula-heavy write row at commit `7118f208`,
> Ascend's focused external repeat-15 run had the fastest median and p95 among
> the completed formula-capable comparable writers in this row. This is scoped
> generated formula-write evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every formula-capable writer."
- "Ascend beats omitted, unsupported, blocked, or untested formula writers."
- "Ascend produces the smallest formula-heavy XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: defer production optimization for formula-heavy from this winning
row. Continue only with a measured release workflow loss, a full-profile
coverage blocker, or an explicit claim downgrade.

## Cycle: Formula Heavy TS/JS/Rust Write Head-to-Head at `2116ddd1`

Classification: comparable external evidence plus benchmark-runner unlock.
The JS writer runners now emit real formula cells for `formula-heavy`, making
SheetJS and ExcelJS ranking-eligible formula-write baselines. Ascend is the
median and p95 winner against SheetJS, ExcelJS, and rust_xlsxwriter in this
focused row. No production optimization is justified from this winning row.

Workflow: generated XLSX write for formula-heavy workbooks, 2000 rows x 20
columns, with 36,000 formula cells and cached values.

Why it matters for release: the user explicitly prioritizes TS/JS head-to-heads
and Rust as the minimum performance floor. This closes the previous JS
capability gap for a formula-bearing generated write workflow and keeps the
result scoped to formula writes rather than broad `xlsx-write-sota` wording.

Public/tracked-clean input: `competitive-io` generated the `formula-heavy`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `2116ddd1`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-formula-js-current-2116ddd1 2116ddd1
cd /private/tmp/ascend-write-formula-js-current-2116ddd1
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-formula-js-current-2116ddd1-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,rust-xlsxwriter --workload formula-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15.json 2> /private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15.json --json --metric p95Ms --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15-p95-scoreboard.json
```

Environment:

- Commit: `2116ddd1c5c2709a7f07b4bedf910783f81da2fe`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-formula-js-current-2116ddd1`; `git status
  --short --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Benchmark payload Node runtime: `24.3.0`; shell `node --version` reported
  `v22.22.0`.
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runner versions: SheetJS `0.18.5`, ExcelJS `4.4.0`, rust_xlsxwriter `0.1.0`.
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload formula-heavy`,
  `validationMode each`, `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15.json
/private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15-time.txt
/private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15-scoreboard.json
/private/tmp/ascend-write-formula-js-current-2116ddd1-runs/write-formula-heavy-js-rust-repeat15-p95-scoreboard.json
```

Formula-heavy JS/Rust writer row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won median and p95 | 14.459 | 19.666 | 0.101 | 203.6 MiB | 405795 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 29.363 | 32.129 | 0.029 | 19.9 MiB | 245490 |
| `sheetjs` | ran/lost vs Ascend | 31.515 | 33.897 | 0.049 | 264.0 MiB | 1844573 |
| `exceljs` | ran/lost vs Ascend | 98.246 | 110.177 | 0.040 | 305.9 MiB | 235051 |

Scoreboard result:

- Median scoreboard: formula-heavy group winner was `ascend-external-writer`;
  `leaderFailures: []` and `profileLeaderFailures: []`.
- P95 scoreboard: formula-heavy group winner was `ascend-external-writer`;
  `leaderFailures: []` and `profileLeaderFailures: []`.
- Full `xlsx-write-sota` coverage still fails, with 61 coverage failures and
  37 coverage gaps, because this row intentionally covers only `formula-heavy`
  and only Ascend, SheetJS, ExcelJS, and rust_xlsxwriter. Missing workloads,
  omitted writers, and unsupported formula writers are not counted as wins.

Semantic comparability: all four rows reopened successfully, matched the
expected one-sheet and 40,000-cell shape, emitted 36,000 formula cells, and
passed sorted semantic cell value validation. Ascend, SheetJS, and ExcelJS
matched ordered semantic cell hashes; rust_xlsxwriter matched sorted semantic
values but not ordered cell hashes. This is formula-write semantic evidence,
not byte-equivalent output evidence. Ascend uses more RSS than rust_xlsxwriter
and emits a larger XLSX than ExcelJS and rust_xlsxwriter.

Humble allowed wording:

> On the generated 2000 x 20 `formula-heavy` write row at commit `2116ddd1`,
> Ascend's external writer was faster by median and p95 than SheetJS `0.18.5`,
> ExcelJS `4.4.0`, and rust_xlsxwriter `0.1.0`, with all rows writing 36,000
> formula cells and passing value validation. This is scoped formula-write
> evidence, not broad `xlsx-write-sota` evidence.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every TS/JS or Rust writer on every workload."
- "Ascend beats omitted, unsupported, blocked, or untested formula writers."
- "Ascend produces the smallest formula-heavy XLSX."
- "Ascend uses less memory than rust_xlsxwriter on formula-heavy writes."
- "Ascend proves byte/order-equivalent output against rust_xlsxwriter."

Next action: defer production optimization from this winning row. Continue with
the next JS/Rust frontier gap: table-heavy JS comparability if the runners can
emit table metadata, feature-rich JS quality boundaries, or a current
full-profile write gate split into attributable JS/Rust workload groups.

## Cycle: Table Heavy Write Current Fastest Comparable Row at `06b2230a`

Classification: comparable external evidence. Ascend is the median and p95
winner among the completed table-capable generated XLSX writers in this row with
acceptable noise. This is a scoped table-write row, not a broad `xlsx-write-sota`
promotion.

Workflow: generated XLSX write for table-heavy workbooks, 2000 rows x 20
columns, including an emitted XLSX table part.

Why it matters for release: table-bearing workbook export is a common
agent-produced report shape. It checks whether Ascend remains competitive on a
release-visible commit/export workflow where the writer must preserve tabular
metadata, not just cell values.

Public/tracked-clean input: `competitive-io` generated the `table-heavy`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `06b2230a`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-table-current-06b2230a 06b2230a8f55
cd /private/tmp/ascend-write-table-current-06b2230a
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-table-current-06b2230a-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,xlsxwriter,openpyxl,rust-xlsxwriter --workload table-heavy --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-table-current-06b2230a-runs/write-table-heavy-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-table-current-06b2230a-runs/write-table-heavy-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-table-current-06b2230a-runs/write-table-heavy-repeat15-scoreboard.json
```

Environment:

- Commit: `06b2230a8f55cee88fc0af8ced101f2fab22923f`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-table-current-06b2230a`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload table-heavy`,
  `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-table-current-06b2230a-runs/write-table-heavy-repeat15.json
/private/tmp/ascend-write-table-current-06b2230a-runs/write-table-heavy-repeat15-scoreboard.json
```

Full comparable table-capable writer row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won | 12.037 | 16.663 | 0.129 | 173.0 MiB | 274212 |
| `rust-xlsxwriter` | ran/lost vs Ascend | 36.493 | 67.496 | 0.277 | 23.4 MiB | 185656 |
| `xlsxwriter` | ran/lost vs Ascend | 142.865 | 175.005 | 0.130 | 109.7 MiB | 188023 |
| `openpyxl` | ran/lost vs Ascend | 236.053 | 333.412 | 0.146 | 105.2 MiB | 152823 |

Scoreboard result:

- Full repeat-15 table-capable row: group winner was
  `ascend-external-writer`; `leaderFailures: []` and
  `profileLeaderFailures: []`.
- The full scoreboard command exits nonzero for full-profile coverage because
  this is not a full `xlsx-write-sota` run. Coverage failures and omitted,
  unsupported, blocked, or untested table writers are not counted as Ascend
  wins.

Semantic comparability: all listed rows reopened successfully, matched the
expected one-sheet and 40,000-cell shape, emitted exactly one table part, and
passed sorted semantic cell value validation. Ascend matched ordered semantic
cell value hashes; XlsxWriter, OpenPyXL, and rust_xlsxwriter did not, so their
rows are useful table-capable write comparisons but not byte/order-equivalent
output claims. Ascend uses more RSS and emits a larger XLSX than the compared
table writers.

Humble allowed wording:

> On the generated 2000 x 20 table-heavy write row at commit `06b2230a`,
> Ascend's focused external repeat-15 run had the fastest median and p95 among
> completed table-capable comparable writers. This is scoped table-write
> evidence, not a broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats every table-capable writer."
- "Ascend beats omitted, unsupported, blocked, or untested table writers."
- "Ascend produces the smallest table-heavy XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: defer production optimization for table-heavy from this current
winning row. Continue only with a measured release workflow loss, a full-profile
coverage blocker, or an explicit claim downgrade.

## Cycle: Feature Rich Write Current Comparable Boundary at `a5fa3006`

Classification: comparable external evidence plus not-comparable boundary.
Ascend is the median and p95 winner against XlsxWriter on the completed
feature-rich generated XLSX write row. OpenPyXL ran but is not ranking eligible
because it failed the rich feature semantic contract, so it is not counted as a
win. This is scoped write-rich-metadata evidence, not a broad `xlsx-write-sota`
promotion.

Workflow: generated XLSX write for feature-rich workbooks, 2000 rows x 20
columns, including defined name, hyperlink, comment, data validation, and
conditional formatting obligations.

Why it matters for release: feature-rich generated export is the closest write
row to a practical agent-authored workbook with visible metadata. Any release
speed claim here must prove both cell values and workbook features, not just
file generation.

Public/tracked-clean input: `competitive-io` generated the `feature-rich`
`source-mode generated-write` workload from tracked benchmark code in a clean
detached worktree at commit `a5fa3006`. No private corpus or local research
workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-feature-current-a5fa3006 a5fa30069de7
cd /private/tmp/ascend-write-feature-current-a5fa3006
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-feature-current-a5fa3006-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,xlsxwriter,openpyxl --workload feature-rich --repeat 15 --warmup 3 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-feature-current-a5fa3006-runs/write-feature-rich-repeat15.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-feature-current-a5fa3006-runs/write-feature-rich-repeat15.json --json --metric medianMs --require-profile xlsx-write-sota --assert-profile-leader ascend > /private/tmp/ascend-write-feature-current-a5fa3006-runs/write-feature-rich-repeat15-scoreboard.json
```

Environment:

- Commit: `a5fa30069de793a6e1b07addebb6b37108968995`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-feature-current-a5fa3006`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category write`, `executionScope external-process`,
  `sourceMode generated-write`, `workload feature-rich`,
  `validationMode each`, `repeat 15`, `warmup 3`.

Raw output:

```text
/private/tmp/ascend-write-feature-current-a5fa3006-runs/write-feature-rich-repeat15.json
/private/tmp/ascend-write-feature-current-a5fa3006-runs/write-feature-rich-repeat15-scoreboard.json
```

Feature-rich writer row, repeat 15 after 3 warmups:

| Runner | Status vs Ascend | Median ms | P95 ms | CV | Peak RSS | Output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ascend-external-writer` | ran/won vs XlsxWriter | 15.270 | 16.568 | 0.070 | 171.0 MiB | 271114 |
| `xlsxwriter` | ran/lost vs Ascend | 160.324 | 232.014 | 0.181 | 104.0 MiB | 121085 |
| `openpyxl` | not comparable | 265.857 | 498.744 | 0.248 | 106.9 MiB | 126574 |

Scoreboard result:

- Feature-rich row: group winner was `ascend-external-writer`;
  `leaderFailures: []` and `profileLeaderFailures: []`.
- OpenPyXL ran but was `semantic-mismatch` and `rankingEligible: false`.
- The scoreboard command exits nonzero for full-profile coverage because this is
  not a full `xlsx-write-sota` run. Coverage failures and omitted, unsupported,
  blocked, semantically mismatched, or untested feature-rich writers are not
  counted as Ascend wins.

Semantic comparability: Ascend and XlsxWriter reopened successfully, matched the
expected one-sheet and 40,000-cell shape, passed sorted semantic cell value
validation, and met the six feature-rich obligations tracked by the scorer.
Ascend matched ordered semantic cell value hashes; XlsxWriter did not, so the
comparison is feature-semantic, not byte/order-equivalent. OpenPyXL reopened and
matched cell values, but it failed the feature-rich semantic contract because
`featureRichSemanticMatches=false` and `featureRichHyperlinkMatches=false`; it
is therefore not comparable for a write-rich-metadata speed claim.

Humble allowed wording:

> On the generated 2000 x 20 feature-rich write row at commit `a5fa3006`,
> Ascend's focused external repeat-15 run was faster than XlsxWriter while
> satisfying the tracked feature-rich obligations. OpenPyXL ran but was not
> semantically comparable, so this is scoped write-rich-metadata evidence, not a
> broad `xlsx-write-sota` claim.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend beats OpenPyXL on feature-rich writes."
- "Ascend beats every feature-rich writer."
- "Ascend beats omitted, unsupported, blocked, semantically mismatched, or
  untested feature-rich writers."
- "Ascend produces the smallest feature-rich XLSX."
- "Ascend proves byte/order-equivalent output against every compared writer."

Next action: defer production optimization for feature-rich from the XlsxWriter
win and keep OpenPyXL as a not-comparable boundary unless the runner or semantic
contract is narrowed in a future claim-specific block.

## Cycle: Feature Rich JS Write Quality Boundary at `9fabfc8e`

Classification: claim boundary. This is not a speed win over JS writers. SheetJS
is explicitly unsupported by the harness for the tracked feature-rich write
contract, and ExcelJS runs but is not ranking eligible because it misses the
tracked comment obligation. Ascend is faster in the emitted timings, but the
release-relevant result is the semantic boundary.

Workflow: generated XLSX write for feature-rich workbooks, 2000 rows x 20
columns, including defined name, hyperlink, comment, data validation, and
conditional formatting obligations.

Why it matters for release: the user prioritizes heads-up JS/TS quality as well
as speed. A feature-rich write claim must not treat value-equivalent output as
equivalent when comments, validations, conditional formatting, or names are part
of the user-visible workbook contract.

Public/tracked-clean input: `competitive-io` generated the `feature-rich`
workload from tracked benchmark code in a clean detached worktree at commit
`9fabfc8e`. No private corpus or local research workbook was used.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-write-feature-exceljs-current-9fabfc8e 9fabfc8e274ce9960c311474227163a082dfd031
cd /private/tmp/ascend-write-feature-exceljs-current-9fabfc8e
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --libraries ascend,sheetjs,exceljs --workload feature-rich --repeat 15 --warmup 3 --validation-mode each > /private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15.json 2> /private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15.json --json --metric medianMs --assert-leader ascend > /private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15-scoreboard.json
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15.json --json --metric p95Ms --assert-leader ascend > /private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15-p95-scoreboard.json
```

Environment:

- Commit: `9fabfc8e274ce9960c311474227163a082dfd031`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-feature-exceljs-current-9fabfc8e`; `git status
  --short --branch` reported `## HEAD (no branch)`.
- Bun runtime: `1.3.13`
- Node: `26.0.0`
- Platform: Darwin arm64, macOS kernel `25.4.0`
- Runner versions: ExcelJS `4.4.0`; SheetJS `0.18.5` was skipped before timing
  because the write runner does not declare `writeRichMetadata`.
- Runtime profile: `category write`, in-process generated writer comparison,
  `workload feature-rich`, `repeat 15`, `warmup 3`, `validationMode each`.

Raw output:

```text
/private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15.json
/private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15-time.txt
/private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15-scoreboard.json
/private/tmp/ascend-write-feature-exceljs-current-9fabfc8e-runs/write-feature-rich-js-repeat15-p95-scoreboard.json
```

Focused JS feature-rich write row, repeat 15 after 3 warmups:

| Runner | Status | Median ms | P95 ms | CV | Peak RSS | Output bytes | Feature result |
| --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| `ascend` | ran/pass | 7.727 | 8.233 | 0.040 | 222.8 MiB | 271114 | all tracked feature obligations pass |
| `exceljs` | ran/semantic-mismatch, not ranking eligible | 63.092 | 65.992 | 0.023 | 348.4 MiB | 123693 | values, order, hyperlink, defined name, validation, and conditional formatting pass; comment obligation fails |
| `sheetjs` | unsupported by harness | n/a | n/a | n/a | n/a | n/a | skipped: runner does not declare `writeRichMetadata=true` |

Process-level `/usr/bin/time -l`: `3.29 real`, `3.72 user`, `0.29 sys`,
`367443968` maximum resident set size, `283853808` peak memory footprint.

Scoreboard result:

- Median scoreboard: feature-rich group winner was `ascend`;
  `leaderFailures: []`, `profileLeaderFailures: []`.
- P95 scoreboard: feature-rich group winner was `ascend`;
  `leaderFailures: []`, `profileLeaderFailures: []`.
- ExcelJS had `correctnessStatus: semantic-mismatch` and
  `rankingEligible: false`, so it is not counted as a speed loss.

Semantic comparability: Ascend and ExcelJS both write the same generated
feature-rich value grid and match ordered semantic cell values. ExcelJS also
matches the tracked hyperlink, defined name, data validation, and conditional
formatting obligations, but `featureRichSemanticMatches=false` because
`featureRichCommentMatches=false`. SheetJS is not attempted for this claim
because the harness does not mark it feature-rich-write capable.

Humble allowed wording:

> On the generated `feature-rich` write row at commit `9fabfc8e`, ExcelJS
> `4.4.0` wrote the values and most tracked metadata but missed the comment
> obligation, so it was marked semantically ineligible. SheetJS was unsupported
> for the tracked rich-metadata write contract. This is a JS feature-rich write
> quality boundary, not a speed win over JS writers.

Forbidden wording:

- "Ascend beats ExcelJS on feature-rich write speed."
- "Ascend beats SheetJS on feature-rich writes."
- "ExcelJS is feature-rich-write comparable for this release claim."
- "SheetJS is feature-rich-write comparable for this release claim."
- "Ascend beats every JS writer on feature-rich writes."

Next action: keep feature-rich write wording honest. Future work should either
improve JS runner semantic coverage, narrow the feature-rich write contract for
a specific user claim, or move to another JS/Rust floor row with comparable
semantics.

## Current Full-Profile Downgrade: XLSX Write SOTA

Classification: blocked claim downgrade. No production optimization target is
justified from this evidence because the attempted broad generated-write gate
did not complete or emit JSON. The completed write rows remain scoped evidence;
they are not a clean full `xlsx-write-sota` promotion artifact.

Workflow: full generated XLSX write coverage for the existing
`xlsx-write-sota` profile, using `competitive-io` generated-write mode and the
existing SOTA writer runner manifest.

Why it matters for release: release wording must not promote "SOTA XLSX write"
from isolated row wins. A full-profile gate needs either complete comparable
coverage or explicit per-runner blockers that do not count as wins.

Public/tracked-clean input: `competitive-io` would generate tracked
`source-mode generated-write` workloads in a clean detached worktree at commit
`eca32509`. No private corpus or local research workbook was used. The run did
not reach sample emission.

Command attempted:

```bash
git worktree add --detach /private/tmp/ascend-write-profile-current-eca32509 eca3250951b3
cd /private/tmp/ascend-write-profile-current-eca32509
bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-write-profile-current-eca32509-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-io.ts --json --category write --competitor all --execution-scope external-process --source-mode generated-write --libraries ascend-external-writer,sheetjs,exceljs,xlsxwriter,xlsxwriter-constant-memory,pyexcelerate,pyexcelerate-range,pyexcelerate-cell,fastexcel-java,openpyxl,openpyxl-write-only,apache-poi,closedxml,npoi,rust-xlsxwriter,excelize --workload all --repeat 5 --warmup 1 --validation-mode each --write-runner-manifest fixtures/benchmarks/runners/sota-writers.manifest.json > /private/tmp/ascend-write-profile-current-eca32509-runs/write-profile-repeat5.json
```

Environment:

- Commit: `eca3250951b3`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-write-profile-current-eca32509`
- Bun install reported Bun `1.3.6` and installed `@types/bun@1.3.13`.
- Command runtime environment used `/Users/arjun/.bun/bin/bun`; the previous
  completed benchmark rows on this machine report Bun runtime `1.3.13`, Node
  `24.3.0`, Darwin arm64.
- Attempt stopped at `2026-05-16T22:36:31Z` after repeated 60-second polls with
  no JSON emitted.

Raw output:

```text
/private/tmp/ascend-write-profile-current-eca32509-runs/write-profile-repeat5.json
```

Result:

| Artifact | Status | Sample count | Median | P95 | CV/noise | Memory/size |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| `write-profile-repeat5.json` | blocked, killed before JSON emission | 0 | n/a | n/a | n/a | n/a |

Operational blocker: the all-workload/all-runner command was too opaque for a
claim-grade gate in this environment. The output file remained zero bytes and no
`competitive-io` process remained after termination. Because no row-level JSON
was produced, this attempt cannot identify a winner, loser, timeout, semantic
mismatch, or optimization target for any individual workflow.

Semantic comparability: none established by this attempted full gate. Existing
row-level sections remain the only valid write evidence: dense-values,
plain-text, sparse-wide, string-heavy, styles-heavy, formula-heavy, table-heavy,
and feature-rich are scoped rows with their own comparability limits.

Humble allowed wording:

> Ascend has scoped external write-row evidence, but the current broad
> `xlsx-write-sota` promotion gate is blocked. The all-workload/all-runner
> generated-write attempt did not emit JSON, so broad XLSX write speed wording
> remains downgraded until the profile is split into attributable workload groups
> or a complete full-profile gate succeeds.

Forbidden wording:

- "Ascend is SOTA for XLSX write."
- "Ascend has passed the full `xlsx-write-sota` gate."
- "Ascend beats every generated XLSX writer."
- "Ascend wins the all-workload writer profile."
- Any wording that counts the killed full-gate attempt, missing rows, omitted
  runners, unsupported semantics, or timeouts as wins.

Next action: split the write profile into existing workload groups and record
per-group coverage, or attack the first completed comparable row that shows a
durable leader failure. Do not do production optimization from this blocked
full-gate attempt.

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

## Cycle: Tracked Real Workbook Strings/Links Open Boundary at `e8654a0b`

Classification: scoped real-workbook evidence plus claim downgrade. Ascend's
SDK value-open surface correctly reads the tracked `strings_links.xlsx` fixture
and is faster than OpenPyXL read-only, Apache POI, and ClosedXML on their
completed lanes, but this row does not support a broad fastest real-workbook
claim. Rust Calamine is faster on its file-path materialization lane, and several
other readers are semantic mismatches on this workbook.

Workflow: real XLSX open/inspect value read for a tracked XlsxWriter workbook
with strings and hyperlinks.

Why it matters for release: this is a real workbook, not a generated synthetic
matrix row. It exercises the first user-visible step of a headless workflow:
open an existing workbook, inspect values, and preserve enough shape information
to reason about the workbook. The row also prevents overclaiming: Ascend has a
correct SDK safe-open result, but a native Calamine reader is faster on a narrower
value materialization lane.

Public/tracked-clean input: `fixtures/xlsx/xlsxwriter/strings_links.xlsx` from
tracked git fixtures in a clean detached worktree at commit
`e8654a0bf7b8689cb203a582b80026fa59dd507c`. The file is 8,801 bytes with SHA256
`e46b7e597607b4d4819ae83265f8d160904e7b01537637db68bff698c46d522b`. No
private corpus or local research workbook was used. A candidate POI workbook was
rejected because it was present only in the dirty local fixture directory and not
tracked in the detached worktree.

Commands:

```bash
git worktree add --detach /private/tmp/ascend-real-workbook-current-e8654a0b e8654a0b
cd /private/tmp/ascend-real-workbook-current-e8654a0b
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun install --frozen-lockfile
mkdir -p /private/tmp/ascend-real-workbook-current-e8654a0b-runs
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-real-workbook.ts --json --category read --competitor external --libraries ascend-external-values,rust-calamine,python-calamine,openpyxl-read-only-values,apache-poi,closedxml,polars-calamine,polars-xlsx2csv,polars-openpyxl,excelize --runner-manifest fixtures/benchmarks/runners/ascend-python-readers.manifest.json --repeat 15 --warmup 3 fixtures/xlsx/xlsxwriter/strings_links.xlsx > /private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-nonordered-repeat15.json 2> /private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-nonordered-repeat15-time.txt
TMPDIR=/private/tmp env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-nonordered-repeat15.json --json --metric medianMs > /private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-nonordered-repeat15-scoreboard-noassert.json
```

Rejected/non-promoted commands:

```bash
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-real-workbook.ts --json --category read --competitor external --libraries ascend-external-values-ordered,rust-calamine,openpyxl-read-only-values,apache-poi,closedxml,polars-calamine,polars-xlsx2csv,polars-openpyxl,excelize --runner-manifest fixtures/benchmarks/runners/nyc311-sota-readers.manifest.json --repeat 15 --warmup 3 fixtures/xlsx/xlsxwriter/styles_formulas.xlsx > /private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-styles-formulas-read-repeat15.json 2> /private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-styles-formulas-read-repeat15-time.txt
TMPDIR=/private/tmp ACCEPT_NPOI_OSMF_LICENSE=1 env PATH=/Users/arjun/.pyenv/shims:/Users/arjun/.bun/bin:/Users/arjun/.cargo/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin /usr/bin/time -l /Users/arjun/.bun/bin/bun run fixtures/benchmarks/competitive-real-workbook.ts --json --category read --competitor external --libraries ascend-external-values-ordered,rust-calamine,openpyxl-read-only-values,apache-poi,closedxml,polars-calamine,polars-xlsx2csv,polars-openpyxl,excelize --runner-manifest fixtures/benchmarks/runners/nyc311-sota-readers.manifest.json --repeat 15 --warmup 3 fixtures/xlsx/xlsxwriter/strings_links.xlsx > /private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-repeat15.json 2> /private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-repeat15-time.txt
```

The `styles_formulas.xlsx` row and the ordered `strings_links.xlsx` row are
not speed evidence because all or several rows were `semantic-mismatch` under
the ordered/hash contract. They are timing-boundary diagnostics only.

Environment:

- Commit: `e8654a0bf7b8689cb203a582b80026fa59dd507c`
- Worktree: clean detached worktree at
  `/private/tmp/ascend-real-workbook-current-e8654a0b`
- Bun runtime: `1.3.13`
- Node: `24.3.0`
- Platform: Darwin arm64
- Runtime profile: `category read`, `competitor external`, `repeat 15`,
  `warmup 3`, tracked file `fixtures/xlsx/xlsxwriter/strings_links.xlsx`.

Raw output:

```text
/private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-nonordered-repeat15.json
/private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-nonordered-repeat15-time.txt
/private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-nonordered-repeat15-scoreboard-noassert.json
/private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-styles-formulas-read-repeat15.json
/private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-repeat15.json
```

Tracked fixture shape:

- Sheet: `Strings`
- Used range: `Strings!A2:D200`
- Cells: 600 logical and physical cells
- Formulas: 0
- Workbook features: 2 worksheet hyperlinks

Repeat 15 after 3 warmups:

| Runner | Status | Median ms | P95 ms | CV | Peak RSS | Timing lane | Semantic comparability |
| --- | --- | ---: | ---: | ---: | ---: | --- | --- |
| `rust-calamine` | ran/won on its lane | 0.650 | 0.759 | 0.073 | 3.0 MiB | `external-internal-file-path-materialization-timing` | Passed shape and value assertions for the 600-cell workbook. Faster and lower RSS than Ascend, but not Ascend's SDK safe-open surface. |
| `ascend-external-values` | ran/won on its lane | 2.162 | 2.813 | 0.109 | 114.7 MiB | `external-internal-file-path-timing` | Passed shape and value assertions through Ascend's SDK value-open path. |
| `openpyxl-read-only-values` | ran/won on its lane, slower than Ascend | 5.367 | 5.727 | 0.040 | 48.9 MiB | `external-internal-read-only-data-only-stream-materialization-timing` | Passed read-only data-only materialization assertions. |
| `closedxml` | ran/won on its lane, slower than Ascend | 12.565 | 22.549 | 0.252 | 112.3 MiB | `external-internal-materialized-workbook-timing` | Passed materialized workbook assertions. |
| `apache-poi` | ran/lost vs ClosedXML, slower than Ascend | 20.145 | 32.402 | 0.185 | 104.0 MiB | `external-internal-materialized-workbook-timing` | Passed materialized workbook assertions. |
| `python-calamine` | not comparable | 0.450 | 0.547 | 0.078 | 28.9 MiB | `external-internal-file-path-materialization-timing` | `semantic-mismatch`; not counted as a win or loss. |
| `polars-calamine` | not comparable | 1.686 | 2.809 | 0.298 | 89.5 MiB | `external-internal-operation-timing` | `semantic-mismatch`; not counted as a win or loss. |
| `polars-xlsx2csv` | not comparable | 5.225 | 9.243 | 0.270 | 63.5 MiB | `external-internal-operation-timing` | `semantic-mismatch`; not counted as a win or loss. |
| `polars-openpyxl` | not comparable | 9.270 | 12.759 | 0.125 | 81.5 MiB | `external-internal-operation-timing` | `semantic-mismatch`; not counted as a win or loss. |
| `excelize` | not comparable | 10.233 | 14.899 | 0.190 | 14.9 MiB | `external-internal-file-path-materialization-timing` | `semantic-mismatch`; not counted as a win or loss. |

Scoreboard result:

- `rust-calamine` was the only eligible row and winner in
  `external-internal-file-path-materialization-timing`.
- `ascend-external-values` was the only eligible row and winner in
  `external-internal-file-path-timing`.
- `openpyxl-read-only-values` was the only eligible row in its read-only stream
  lane.
- `closedxml` beat `apache-poi` in `external-internal-materialized-workbook-timing`.
- Polars, Python Calamine, and Excelize rows were semantic mismatches and are
  not counted as wins or losses.

Semantic boundary: this is tracked real-workbook evidence for value open/inspect
only. It does not prove formula calculation, style fidelity, package
preservation, hyperlink preservation after save, or broad real-workbook speed.
The timing lanes are not one unified timing boundary, so do not collapse them
into a single cross-library leaderboard. The row is still a concrete downgrade
for any broad wording that would imply Ascend is fastest on real-workbook
open/inspect; Rust Calamine is faster on the narrower materialization lane.

Humble allowed wording:

> On the tracked `strings_links.xlsx` real workbook at commit `e8654a0b`,
> Ascend's SDK value-open path passed the 600-cell shape/value contract and was
> faster than completed OpenPyXL, Apache POI, and ClosedXML rows. Rust Calamine
> was faster and lower-memory on its narrower materialization lane, while several
> other readers were semantic mismatches. This is scoped real-workbook
> open/inspect evidence, not a broad speed claim.

Forbidden wording:

- "Ascend is fastest for real-workbook open/inspect."
- "Ascend beats Calamine on real workbooks."
- "Ascend beats every reader on `strings_links.xlsx`."
- "Ascend beats Python Calamine, Polars, or Excelize on this row."
- Any wording that treats semantic mismatches or different timing lanes as wins.

Next action: keep broad real-workbook speed wording downgraded. Continue with a
larger tracked fixture or a claim-specific same-lane runner only if it can
produce comparable semantics without using unapproved local research workbooks.

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
