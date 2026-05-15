# Blocked Commit Rollback Proof

## Question

When `commitAgentPlanFromWorkbook` blocks a write after applying operations for pre-write policy checks, does it restore the caller's in-memory workbook before throwing?

## Hypothesis

It should. Agent commit failures must be fail-closed not only for output files, but also for the workbook object the agent is holding. A blocked write-policy diagnostic should not leave speculative edits in memory.

## External sources checked

- Language Server Protocol diagnostics model rejected edits as structured diagnostics rather than partially applied workspace changes: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- OpenTelemetry semantic conventions reinforce that failure diagnostics should be structured and explainable for downstream tooling: https://opentelemetry.io/docs/specs/otel/semantic-conventions/
- SLSA provenance distinguishes local evidence from verifiable release provenance; this fix strengthens local agent evidence but does not create provenance: https://slsa.dev/spec/v1.0-rc1/provenance

## Why this matters to Ascend

The auditable mutation claim depends on trustworthy plan/commit behavior. If a blocked commit mutates the in-memory workbook, an agent can continue from a state that was never written or approved. That competes directly with preservation-first, fail-closed mutation planning.

## Probe/implementation

The probe started as a one-line assertion in the real imported shared-formula trust moat:

```ts
expect(wb.sheet('Label')?.cell('C1')).toBeUndefined()
```

It initially failed: `commitAgentPlanFromWorkbook` did not write the output file, but the speculative `setCells` operation remained in the in-memory workbook after the write-policy blocker.

Fold-in:

- added an SDK workbook mutation rollback snapshot that captures workbook model, dirty flags, package graph cache, pending dirty refs, and generation counters;
- wrapped the commit apply/recalc/preservation/write-policy phase in a rollback-on-error path;
- kept successful commits unchanged;
- retained the real shared-formula diagnostic ref assertions from the prior proof.

Commands run:

```bash
bun test packages/sdk/src/agent-workflow.test.ts
bunx biome check packages/sdk/src/agent-workflow.ts packages/sdk/src/workbook.ts packages/sdk/src/agent-workflow.test.ts
bunx tsc --build
```

## Results

- The regression assertion now passes: a blocked real shared-formula commit leaves `Label!C1` absent in the source workbook.
- SDK agent workflow tests passed: 74 tests, 498 assertions.
- Biome passed for the touched SDK files.
- `bunx tsc --build` passed.

## Confidence

High for the write-policy blocker path covered by the real shared-formula fixture. Medium for all possible post-apply failure paths until broader rollback tests cover recalc failure, write failure, and post-write verification failure.

## Fold-in decision

Promote to correctness loop. This is a production correctness fix with a focused regression test. It does not add a product surface or relax release claim gates.

## Next question

Should agent commit rollback semantics be generalized into an explicit invariant test matrix for apply failure, recalc failure, write-policy blocker, write failure, and post-write verification failure?
