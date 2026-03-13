---
name: autooptimize
description: Autonomous optimization loop for benchmarks and performance. Use when optimizing a benchmark scenario, improving performance, reducing latency, lowering memory usage, or when the user says "autooptimize", "optimize benchmark", "make it faster", "performance loop", or "run autoopt".
---

# Autooptimize

Autonomous experiment loop inspired by [karpathy/autoresearch](https://github.com/karpathy/autoresearch). You make small, focused code changes, measure, keep what works, revert what doesn't, repeat. Never stop unless interrupted.

Read AGENTS.md for project structure, commands, commit conventions, and benchmark infrastructure.

## Setup

When the user triggers autooptimize, do the following once:

### 1. Determine the target

Infer from the user's request, or ask. You need:

| Field | Example | Required |
|-------|---------|----------|
| **target** | `recalc-incremental` | yes |
| **type** | `bench-scenario`, `micro`, `test-speed`, `custom` | yes |
| **metric** | `medianMs`, `opsPerSec`, wall time | yes |
| **direction** | `lower` or `higher` | yes |
| **scope** | `packages/engine/src/` | yes |

Run `bun bench` to list all available benchmark scenarios. Run `bun run fixtures/benchmarks/micro.ts` for microbenchmarks. See AGENTS.md for full benchmark infrastructure details.

### 2. Create branch and session directory

```bash
git checkout -b autoopt/<target>
mkdir -p .autoopt
```

### 3. Read the code in scope

Thoroughly read every file in the scope directory. Understand the data structures, hot paths, allocation patterns, and algorithmic complexity. This is the most important step — you can't optimize what you don't understand.

### 4. Run baseline measurement

For `bench-scenario`:
```bash
bun bench --scenario <target> --repeat 5 --json
```

For `micro`:
```bash
bun run fixtures/benchmarks/micro.ts --json
```

For `test-speed`:
```bash
time bun test --recursive
```

### 5. Write session files

Create `.autoopt/session.md` with this template:

```markdown
# Autooptimize: <target>

## Objective
<one sentence: what are we optimizing and why>

## Target
- **Scenario**: <target name>
- **Type**: <bench-scenario | micro | test-speed | custom>
- **Metric**: <medianMs | opsPerSec | wall time>
- **Direction**: <lower | higher>
- **Baseline**: <value with unit>
- **Current best**: <value with unit> (run N, delta%)

## Scope
<list every file the agent may modify>

## Key Wins
<numbered list of experiments that improved the metric, with description and delta>

## Dead Ends
<numbered list of approaches that didn't work, with brief explanation of why>

## Next Ideas
- [ ] <concrete optimization idea>
- [ ] <concrete optimization idea>
```

Create `.autoopt/results.jsonl` with the config line:

```jsonl
{"type":"config","target":"<name>","metric":"<metric>","direction":"<lower|higher>","command":"<full benchmark command>","scope":["<path>"],"baseline":<number>,"timestamp":<epoch ms>}
```

Then append the baseline result:

```jsonl
{"run":1,"metric":<value>,"status":"baseline","commit":"<sha>","description":"baseline measurement","timestamp":<epoch ms>}
```

## The Loop

Run this loop. Never stop unless interrupted.

```
1. ANALYZE   — read code in scope, identify one optimization opportunity
2. HYPOTHESIZE — describe what you'll change and why it should help
3. EDIT      — make the change (small, focused, one idea per experiment)
4. VERIFY    — run tests: bun test --recursive
               if tests fail → REVERT → log "fail" → go to 1
5. MEASURE   — run the benchmark command with --repeat 5 --json
               parse the metric from JSON output
6. DECIDE    — compare metric to current best (see Decision Rules)
               KEEP: git add -A && git commit → log "keep" → update session
               DISCARD: git checkout -- . → log "discard" → update session
7. REFLECT   — update Next Ideas in session.md, remove tried ideas
8. GO TO 1
```

**Pacing**: each iteration should take 2-5 minutes. If you're spending more than 10 minutes on a single idea, it's too complex — break it down or skip it.

**Commit messages** for kept experiments follow AGENTS.md conventions with `perf(<scope>):` prefix:
```
perf(<scope>): <what you did>

Autooptimize run <N>: <target> <baseline> → <new value> (<delta>%)
```

## Measurement

### Benchmark scenarios

```bash
bun bench --scenario <target> --repeat 5 --json 2>/dev/null
```

Parse the JSON output. The metric is at `.metrics.medianMs` (lower is better). Also note `.metrics.throughputPerSec` and `.metrics.heapDeltaBytes` as secondary metrics.

If the improvement is borderline (2-5%), re-measure with `--repeat 10` to confirm:
```bash
bun bench --scenario <target> --repeat 10 --json 2>/dev/null
```

### Microbenchmarks

```bash
bun run fixtures/benchmarks/micro.ts --json 2>/dev/null
```

Output is an array. Find the matching benchmark by name. Metric is `.opsPerSec` (higher is better).

### Test speed

```bash
/usr/bin/time -p bun test --recursive 2>&1
```

Metric is wall time from the output (lower is better).

### Secondary metrics

Always check that secondary metrics don't regress badly:
- Memory (heapDeltaBytes, rssDeltaBytes) should not increase by more than 50%
- Throughput should not decrease if you're optimizing latency
- Test count should not decrease

## Decision Rules

| Condition | Decision |
|-----------|----------|
| Primary metric improved >5% | **KEEP** |
| Improved 2-5% | Re-measure with `--repeat 10`. Keep if confirmed. |
| Improved <2% | **DISCARD** (likely noise) |
| Unchanged | **DISCARD** |
| Regressed | **DISCARD** |
| Equal perf but simpler/cleaner code | **KEEP** (simplification win) |
| Memory regressed >50% for a small speed gain | **DISCARD** |
| Same idea discarded twice | **SKIP** permanently, note in Dead Ends |

## Results Log Format

Each experiment appends one line to `.autoopt/results.jsonl`:

```jsonl
{"run":2,"metric":0.38,"delta":"-7.3%","status":"keep","commit":"abc123","description":"replaced Map.forEach with for-of in evaluateCompiled","timestamp":1710000002000}
{"run":3,"metric":0.43,"delta":"+4.9%","status":"discard","description":"tried SharedArrayBuffer for cell values — copy overhead exceeded benefit","timestamp":1710000003000}
{"run":4,"metric":0.41,"delta":"0%","status":"fail","description":"inlined resolveRef — test recalc-cross-sheet broke","timestamp":1710000004000}
```

Fields: `run` (sequential), `metric` (measured value), `delta` (% vs current best), `status` (baseline|keep|discard|fail), `commit` (SHA if kept, omit if discarded), `description` (what you tried, concise).

## Resume Protocol

If you find existing `.autoopt/session.md` and `.autoopt/results.jsonl`:

1. Read both files completely
2. Identify the current best metric and run number
3. Read the Dead Ends to avoid repeating failed approaches
4. Read Next Ideas for the backlog
5. Continue the loop from the next run number

This lets a fresh agent context resume exactly where the previous one left off.

## Optimization Patterns

Patterns to look for in this codebase, roughly ordered by typical impact:

**Algorithmic**
- Replace O(n) scans with indexed lookups or binary search
- Cache computed values that are re-derived on every call
- Batch operations instead of one-at-a-time
- Skip unnecessary work (early returns, dirty flags, incremental updates)

**Allocation**
- Reuse arrays/objects instead of allocating in hot loops
- Replace `.map()/.filter()/.reduce()` chains with index loops on hot paths
- Replace `structuredClone` with manual property copy
- Pool temporary objects

**V8/JSC**
- Maintain monomorphic object shapes (always set all properties, same order)
- Avoid megamorphic property access (too many shapes at one call site)
- Use `| 0` for integer arithmetic in hot loops
- Prefer `for` loops over `for...of` for arrays in hot paths

**Data structures**
- `Float64Array` / `Int32Array` for homogeneous numeric data
- Plain objects vs `Map` — benchmark both, perf depends on usage pattern
- Flat arrays vs nested objects for cache-friendly access
- Consider WASM for tight numeric loops (project already has `wasm-range.ts` precedent)

**Codebase-specific**
- SparseGrid: row-major storage, batch set operations, typed value storage
- DependencyGraph: interval index queries, dirty tracking granularity
- Compiled eval: opcode dispatch, stack sizing, range operation fusion
- Formula parse: cache hit rates, AST node allocation
- Style registry: hash-based dedup, style interning

## Rules

- **Never break tests.** Correctness is non-negotiable. If tests fail, revert immediately.
- **One idea per experiment.** Don't combine changes — you can't tell what worked.
- **Small diffs.** Each experiment should change at most 2-3 files, ideally 1.
- **Never modify test files** to make tests pass. Fix the implementation.
- **Never modify benchmark files.** The measurement harness is fixed.
- **Never commit `.autoopt/`** session files. They're ephemeral agent context.
- **Log everything.** Every experiment gets a results.jsonl entry, even failures.
- **Don't chase noise.** Sub-2% changes are indistinguishable from measurement variance.
- **Know when to stop a direction.** If 3 variations of the same idea all fail, it's a dead end. Move on.
- **Respect the codebase conventions.** Pure functions, no side effects, no ambient state. Read AGENTS.md.
