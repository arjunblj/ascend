# Agent Phase Progress Timeout Guard

## Question

Can the agent phase-profile benchmark support long-running performance-owner probes with machine-readable progress and a fail-closed timeout, without promoting latency claims?

## Hypothesis

Yes. The existing workflow progress events can be streamed as NDJSON on stderr while the benchmark keeps the final JSON result on stdout. A timeout guard can emit a compact timeout payload and exit nonzero if a local run stalls.

## External sources checked

- Bun child process docs show `stdout` and `stderr` can be piped independently from `Bun.spawn`: https://bun.com/docs/api/spawn
- Bun's stderr guide documents reading child-process stderr as a stream: https://bun.sh/docs/guides/process/spawn-stderr
- Node timers documentation defines `setTimeout`/`clearTimeout` behavior for fail-closed guards: https://nodejs.org/api/timers.html

## Why this matters to Ascend

Performance owners need practical ways to run edit/verify probes on larger public inputs without losing visibility while a run is in progress. Streaming progress also helps distinguish a slow phase from a stalled harness, but it must remain benchmark evidence rather than release performance wording.

## Probe/implementation

Finished the in-flight `fixtures/benchmarks/agent-phase-profile.ts` harness change:

- added `--progress` to stream phase timings as JSON lines on stderr;
- added `--timeout-ms` to emit a timeout JSON payload and exit `124` if the run stalls;
- kept the final benchmark JSON on stdout;
- tightened invalid timeout parsing to fall back to five minutes instead of an immediate zero-delay timeout;
- added regression coverage for progress streaming and timeout guard parsing in `fixtures/benchmarks/agent-phase-profile.test.ts`.

Commands run:

```bash
bun test fixtures/benchmarks/agent-phase-profile.test.ts
bunx biome check fixtures/benchmarks/agent-phase-profile.ts fixtures/benchmarks/agent-phase-profile.test.ts
bun run fixtures/benchmarks/agent-phase-profile.ts --rows 40 --cols 4 --updates 2 --repeat 1 --warmup 0 --timeout-ms 300000 --progress --json
```

## Results

- Focused benchmark tests passed: 3 tests, 44 assertions.
- Biome passed for the touched benchmark files.
- Direct smoke probe wrote final JSON to stdout and 65 progress JSON lines to stderr.
- The progress stream included both normal and prepared commit `write` and `post-write` phase events for sample 1.

## Confidence

High for progress event streaming and timeout parsing. Medium for owner-loop usefulness until a performance owner runs this against approved public inputs and decides whether progress output should be consumed by release report tooling.

## Fold-in decision

Fold into the performance benchmark harness. Keep practical latency and phase-progress evidence excluded from release proof claims until release-environment runs and threshold wording are owner-approved.

## Next question

Can the performance owner consume progress NDJSON in CI or local scripts to capture partial evidence when long public-input probes time out, without turning timeout output into a headline claim?
