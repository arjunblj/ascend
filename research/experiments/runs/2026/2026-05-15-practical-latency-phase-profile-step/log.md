# Practical Latency Phase Profile Step

## Question

Can the practical latency contract expose an edit-verify phase-profile step so performance owners can identify plan/commit bottlenecks without turning dirty or local-input runs into release claims?

## Hypothesis

Yes. Adding the existing `agent-phase-profile` runner to the edit-verify contract should make load-workbook, preview, preservation audit, dirty write, and reopen phases visible in the report while preserving the current worktree and input-provenance guardrails.

## External sources checked

- Bun documents CPU and heap profiling flags and markdown profile output for local bottleneck diagnosis: https://bun.sh/docs/project/benchmarking
- Google Benchmark's guide emphasizes explicit benchmark output, fixture setup, and reporting controls: https://google.github.io/benchmark/user_guide.html
- hyperfine documents warmup runs and preparation commands as benchmark hygiene: https://github.com/sharkdp/hyperfine

## Why this matters to Ascend

The top claims, safe unknown workbook opening and auditable package-part mutation, both depend on honest latency reporting. If a single edit-verify total is slow, a performance owner needs phase evidence before changing code. The claim board also needs a fail-closed boundary: diagnostic local/private benchmark runs must not become release performance wording.

## Probe/implementation

Commit `ac3fcf8c test(benchmarks): split edit contract phases` folded the existing `fixtures/benchmarks/agent-phase-profile.ts` runner into `fixtures/benchmarks/practical-latency-contracts.ts`.

The practical latency report now includes:

- an `agent-phase-profile` edit-verify step;
- a profile command labelled `contract-edit-verify-phase-profile`;
- diagnostic rows for shared plan load-workbook, preview, preservation audit, shared commit dirty write, and shared commit reopen output;
- phase candidates in the production target decision matrix when real phase timings exist.

Probe command:

```bash
bun run fixtures/benchmarks/practical-latency-contracts.ts --dry-run --contract edit-verify --repeat 1 --warmup 0 --json --out-dir /tmp/ascend-phase-profile-dry-run
```

## Results

The dry run emitted three skipped edit-verify steps:

- `workflow-commit`
- `post-write-breakdown`
- `agent-phase-profile`

The new phase step command was:

```bash
bun run fixtures/benchmarks/agent-phase-profile.ts --input-file fixtures/xlsx/stress/dense-100k.xlsx --updates 25 --repeat 1 --warmup 0 --json
```

The profile command was:

```bash
bun run fixtures/benchmarks/profile-bun.ts --mode all-md --label contract-edit-verify-phase-profile --out-dir /tmp/ascend-phase-profile-dry-run/profiles -- bun run fixtures/benchmarks/agent-phase-profile.ts --input-file fixtures/xlsx/stress/dense-100k.xlsx --updates 25 --repeat 1 --warmup 0 --json
```

The report stayed fail-closed: the run was diagnostic only because the benchmark inputs were local/private and the worktree had unrelated tracked engine edits. No latency threshold or release claim was promoted.

Validation already completed for the implementation checkpoint:

```bash
bunx biome check fixtures/benchmarks/practical-latency-contracts.ts
```

## Confidence

High that the report shape and guardrail work. Medium for performance ownership because this dry run proves wiring only; it intentionally does not produce timings.

## Fold-in decision

Folded into the performance benchmark harness. Do not promote practical latency reports into the release proof index until a tracked-clean run over standardized public inputs has owner-approved threshold wording.

## Next question

Can the current claim board be refreshed from proof artifacts while keeping practical latency evidence explicitly excluded from headline release claims?
