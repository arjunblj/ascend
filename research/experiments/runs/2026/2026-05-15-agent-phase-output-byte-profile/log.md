# Agent Phase Output Byte Profile

## Question

Can the agent phase profile report actual output workbook bytes alongside payload bytes, so performance owners can separate JSON response size from written XLSX size?

## Hypothesis

Yes. The benchmark already writes normal and prepared commit outputs. Recording `stat()` sizes after each write adds a low-cost measurement that makes edit/verify profiles more auditable without changing production behavior.

## External sources checked

- Bun benchmarking docs recommend using `performance.now()` for timing and describe profiling/benchmarking workflows for Bun scripts: https://bun.com/docs/project/benchmarking
- hyperfine documents structured export formats and benchmark metadata for later analysis, supporting the principle that benchmark reports should include machine-readable measurements rather than relying on prose: https://github.com/sharkdp/hyperfine
- SLSA provenance separates build inputs/outputs and downstream verification expectations; output workbook size is useful local evidence but not signed provenance: https://slsa.dev/spec/v1.0-rc1/provenance

## Why this matters to Ascend

The performance loop needs to know whether an edit/verify optimization changes the user-visible commit envelope, the JSON payload shape, or the actual XLSX output size. Without output bytes, the phase profile can show payload bytes but not whether normal and prepared commit paths produce comparable workbook artifacts.

## Probe/implementation

Finished the in-flight `fixtures/benchmarks/agent-phase-profile.ts` change:

- imported `stat` from `node:fs/promises`;
- measured `commitOutputBytes` after the normal commit writes output;
- measured `sharedCommitOutputBytes` after the prepared/shared commit writes output;
- summarized both as medians;
- extended `fixtures/benchmarks/agent-phase-profile.test.ts` to assert both medians for generated and existing-input profiles.

Commands run:

```bash
bun test fixtures/benchmarks/agent-phase-profile.test.ts
bun run fixtures/benchmarks/agent-phase-profile.ts --rows 80 --cols 6 --updates 12 --repeat 1 --warmup 0 --json
```

## Results

- Focused test passed: 2 tests, 38 assertions.
- Local small generated probe reported:
  - `commitOutputBytesMedian=5802`;
  - `sharedCommitOutputBytesMedian=5802`;
  - `postWriteValid=true`;
  - `sharedPostWriteValid=true`.
- The measurement is benchmark-only and does not promote a release performance claim.

## Confidence

High for the measurement plumbing and test coverage. Medium for release usefulness until a performance owner runs the full public-tracked contract on a clean release environment.

## Fold-in decision

Fold into the performance benchmark harness. Keep practical latency excluded from release proof until tracked-clean public-input runs and threshold wording are owner-approved.

## Next question

Can practical latency contract summaries surface output bytes from `agent-phase-profile` in the owner-facing report without implying a release threshold?
