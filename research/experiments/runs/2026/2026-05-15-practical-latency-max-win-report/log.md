# Practical Latency Max Win Report

## Question

Can practical latency reports distinguish phase share from maximum plausible user-visible win so performance owners do not overread diagnostic timings?

## Hypothesis

Yes. The decision matrix should show both max plausible win milliseconds and percent, clamped to the user-visible envelope, instead of implying every phase median is automatically realizable as an optimization win.

## External sources checked

- Google Benchmark emphasizes benchmark reports as explicit measurement artifacts rather than product claims: https://google.github.io/benchmark/user_guide.html
- hyperfine documents warmups and benchmark hygiene for command-line timing evidence: https://github.com/sharkdp/hyperfine
- Bun documents local benchmarking/profiling support used by the practical latency harness: https://bun.sh/docs/project/benchmarking

## Why this matters to Ascend

Performance research should point owners at one concrete target without turning diagnostic ceilings into release claims. A clamped max-win column makes the product boundary clearer: optimize the largest user-visible envelope, not just the largest internal phase number.

## Probe/implementation

Commit `4ea7f427 test(benchmarks): report max plausible contract wins` updated `fixtures/benchmarks/practical-latency-contracts.ts`:

- `EnvelopeDecision` now includes `maxPlausibleWinMs` and `maxPlausibleWinPct`.
- `productionTarget()` reports both values.
- The decision matrix has separate "Max plausible win ms" and "Max plausible win %" columns.
- `maxPlausibleWin()` clamps phase time to the envelope.

## Results

Validation:

```bash
bun run fixtures/benchmarks/practical-latency-contracts.ts --dry-run --contract edit-verify --repeat 1 --warmup 0 --json --out-dir /tmp/ascend-phase-profile-dry-run-2
bunx biome check fixtures/benchmarks/practical-latency-contracts.ts
bunx tsc --build
bun run test:changed
```

Results:

- dry run emitted the edit-verify workflow, post-write, and agent-phase-profile steps;
- dry run stayed diagnostic-only because the benchmark inputs are local/private and research docs were dirty;
- Biome and typecheck passed;
- changed tests passed with 4379 pass, 1 skip, 0 fail across 168 files.

## Confidence

High for report semantics and wiring. Low for any performance claim from this run because it intentionally produced no timings and used local/private benchmark inputs.

## Fold-in decision

Folded into the performance harness only. Practical latency reports remain excluded from release proof until a tracked-clean, public-input, owner-approved run exists.

## Next question

Can the release claim board keep practical latency in the "excluded evidence" lane while top claim owners work only on safe-open and package-action proof gates?
