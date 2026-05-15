# Agent First Window Progress Timeout

## Question

Should long first-window benchmark probes emit progress events and support a timeout guard so autonomous loops do not sit silently during slow UI/API/MCP measurements?

## Hypothesis

Yes. Progress events on stderr and an explicit timeout make performance-owner probes more observable without changing benchmark JSON stdout or product surfaces.

## External sources checked

- Bun benchmarking documentation: https://bun.com/docs/project/benchmarking
- Google Benchmark user guide: https://google.github.io/benchmark/user_guide.html
- web.dev performance budgets overview: https://web.dev/articles/performance-budgets-101
- MDN performance budgets: https://developer.mozilla.org/en-US/docs/Web/Performance/Performance_budgets

## Why this matters to Ascend

The North Star includes real-world performance and agent DX. First-window probes cover SDK, API, CLI, MCP, and TUI paths; when those probes are slow or stuck, owner loops need partial progress evidence without corrupting machine-readable benchmark results.

## Probe/implementation

Finished the in-flight `fixtures/benchmarks/agent-first-window.ts` harness change:

- added `--progress` JSONL events on stderr;
- added `--timeout-ms` elapsed-time guard;
- emits input, warmup, case, sample, heartbeat, and completion events while keeping JSON output on stdout;
- clears the heartbeat interval in `finally`;
- added a focused test proving progress events do not pollute JSON output.

## Results

Focused validation passed:

```bash
bun test fixtures/benchmarks/agent-first-window.test.ts -t "progress events|isolate one"
bunx biome check fixtures/benchmarks/agent-first-window.ts fixtures/benchmarks/agent-first-window.test.ts
bunx tsc --build
```

The progress test verifies that `--progress --json --only capped` returns parseable JSON on stdout and `agent-first-window` events including `input-ready` and `sample-complete` on stderr.

## Confidence

High that the harness now provides useful progress without breaking JSON consumers. Medium that the timeout phase labels are complete for every slow sub-call, because timeout checks are at warmup/sample boundaries rather than inside each lower-level read path.

## Fold-in decision

Promote to the performance benchmark harness. Keep this as owner-routing observability, not a release latency claim and not a product API.

## Next question

Can first-window and practical-latency progress events share a small event schema so long-running performance probes are easier to compare without adding a public surface?
