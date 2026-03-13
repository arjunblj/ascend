# Ascend 2-Day Cloud Sprint Report

## 1) What changed

### Landed optimizations

1. **SparseGrid scan optimization**
   - File: `packages/core/src/sparse-grid.ts`
   - Change: cached sorted chunk-row and chunk-col key lists, invalidated only on structural map changes.
   - Goal: reduce repeated sorting overhead in `iterateRows()` and `iterateRowsInRange()` hot paths used by analysis and SDK window reads.

### Benchmark additions

1. **Microbenchmarks (`fixtures/benchmarks/micro.ts`)**
   - `SparseGrid.iterateRows (40k cells)`
   - `SparseGrid.iterateRowsInRange (window scans)`
   - `SparseGrid structural row/col edits`
   - Evaluator shootout:
     - `Evaluator tree eval (10k formulas)`
     - `Evaluator codegen eval (10k formulas)`
     - `Evaluator SUM range path (2k formulas)`
   - Recalc phase attribution:
     - `Recalc phase: analysis (3k formulas)`
     - `Recalc phase: dependency graph build (3k formulas)`
     - `Recalc phase: dirty-set (1k formulas)`
     - `Recalc phase: eval-order (1k formulas)`

2. **Scenario benchmarks (`fixtures/benchmarks/run.ts`)**
   - `workflow-first-window-cold`
   - `workflow-style-heavy-roundtrip`
   - `recalc-phase-analysis`
   - `recalc-phase-graph-build`
   - `recalc-phase-dirty-set`
   - `recalc-phase-eval-order`
   - `recalc-overlapping-range-deps`
   - New sets:
     - `canary` (always-on PR gating set)
     - `observe` (non-gating noisy/heavier set)
   - Expanded `smoke` set with phase/first-window/style-heavy coverage.

3. **Real workbook benchmarks (`fixtures/benchmarks/real-workbook.ts`)**
   - Added `workflow-session-recalc` step for real-workbook recalc coverage.

### CI / policy hardening

1. **Benchmark workflow (`.github/workflows/bench.yml`)**
   - Pinned Bun to `1.3.10` (removed `latest` drift).
   - Added always-on canary run for PRs.
   - Added observe-only benchmark run (`continue-on-error`).
   - Tightened comparison thresholds (`median=12%`, `retained RSS=25%`).
   - Added comparison artifact upload and clearer summary output.

2. **Target policy (`fixtures/benchmarks/targets.ts`)**
   - Replaced “single best scenario wins category” with median-throughput evaluation across matching scenarios.

### SDK guidance note

- Added `docs/sdk-next-sprint-direction.md`
  - Preserves `preview()/review()` center of gravity.
  - Preserves explicit partial-view capability errors.
  - Documents `open -> select -> propose -> review -> commit -> save`.
  - Keeps `Operation[]` as advanced escape hatch.

## 2) What improved

Primary kept optimization evidence (baseline repeat=5 vs post-change repeat=5):

- `recalc-defined-names-heavy`
  - median: `20.656 ms -> 18.457 ms` (**-10.64%**)
  - throughput: `290,477 -> 325,075` (**+11.91%**)
- `read-window-dense-values`
  - median: `50.716 ms -> 49.662 ms` (**-2.08%**)
  - throughput: `1,971,768 -> 2,013,610` (**+2.12%**)
- `recalc-incremental`
  - median: `0.518 ms -> 0.511 ms` (**-1.39%**, near-noise)

Other tracked baseline scenarios (same command class, repeat=5):

- `recalc-sumifs-large`: `-3.81%` median
- `recalc-1m-dense`: `-14.39%` median (large gain, but should be re-confirmed in CI due variance sensitivity)
- `recalc-lookup-exact-incremental`: `+3.79%` median (possible regression/noise; needs follow-up with higher sample count on stable runner)

Memory notes:

- No material heap growth in most scenarios.
- `recalc-defined-names-heavy` retained RSS increased from `2.70 MB` to `4.87 MB` in this sample set; speedup is strong, but this should be confirmed on CI hardware before declaring a memory-neutral win.

## 3) What was tried and rejected

1. **Initial evaluator codegen import path**
   - Attempted import through engine barrel.
   - Rejected because `codegenFormula` is not exported there; switched to direct module import.

No speculative subsystem rewrites were introduced.

## 4) Current best understanding

1. `SparseGrid` repeated ordering work was a real hotspot in scan-heavy paths.
2. Recalc phase attribution now exists in both micro and scenario suites, so analysis/graph/dirty/eval-order costs are measurable independently.
3. Remaining bottlenecks likely include:
   - evaluator/range materialization behavior on lookup/range-heavy formulas,
   - dirty propagation/eval-order overhead in large dependency surfaces,
   - style-heavy save/apply pipelines under repeated edits.

## 5) Next three best experiments

1. **Evaluator range materialization reduction**
   - Benchmark driver: `Evaluator SUM range path`, `recalc-overlapping-range-deps`, `recalc-lookup-exact-incremental`.
   - Experiment: stream/reuse range buffers in evaluator hot functions, avoid per-call array materialization.

2. **Dirty-set and eval-order narrowing**
   - Benchmark driver: `recalc-phase-dirty-set`, `recalc-phase-eval-order`, `recalc-incremental`.
   - Experiment: tighten dirty seed expansion and reduce repeated set/index rebuilds for small dirty edits.

3. **Style-heavy workflow memory discipline**
   - Benchmark driver: `workflow-style-heavy-roundtrip`, real-workbook `workflow-session-recalc`.
   - Experiment: reuse temporary structures across repeated formatting operations; verify retained RSS behavior on canary and observe sets.

## Validation summary

- `bunx tsc --build`: pass
- `bunx biome check`: pass
- `bun fixtures/benchmarks/run.ts --set smoke --repeat 3 --json`: pass
- `bun run fixtures/benchmarks/micro.ts --json`: pass
- `bun test --recursive`: 4 known pre-existing failures (unchanged from baseline):
  - CLI inspect pivots/slicers/drawings tests
  - `WorkbookReadView` / `WorkbookDocument` pivot+slicer surface test
