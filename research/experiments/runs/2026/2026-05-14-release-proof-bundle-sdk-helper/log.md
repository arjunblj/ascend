# Release Proof Bundle SDK Helper

## Question

Can the release proof bundle experiment be folded into a small SDK surface that packages existing plan, commit, reopen, diff, and audit evidence without inventing new claims?

## Hypothesis

Yes. The prior release proof probe showed that `createAgentPlan` and `commitAgentPlan` already produce most of the evidence. A production helper can make that evidence durable by linking input hashes, plan digest, operation artifact digests, trace digests, post-write reopen results, package graph audits, optional diff evidence, and explicit claim boundaries.

## External sources checked

- [SLSA Provenance](https://slsa.dev/spec/v1.0-rc1/provenance): release evidence should identify the subject artifact, input materials, invocation, and output digests.
- [SLSA attestation model](https://slsa.dev/spec/v1.0/attestation-model): metadata should be explicit; a signature by itself does not define the claim.
- [GitHub artifact attestations](https://docs.github.com/actions/security-guides/using-artifact-attestations-to-establish-provenance-for-builds): signed artifact provenance can be verified separately from the artifact, but that requires a real attestation workflow.
- [in-toto attestation framework](https://github.com/in-toto/attestation): attestation statements bind subjects to typed predicates and metadata.

## Why this matters to Ascend

Trustworthy mutation planning only becomes useful if agents and humans can carry the proof around. A release proof bundle gives future CLI/API/MCP surfaces a stable object to show: what input was inspected, what plan was committed, what output was written, whether post-write reopen and audits passed, and what the proof does not claim.

## Probe/implementation

Inspected local implementation:

- `packages/sdk/src/agent-workflow.ts` already exposes plan and commit results with trace/artifact digests, package graph audits, write policy diagnostics, post-write verification, input/output hashes, and compact summaries.
- `packages/sdk/src/agent-workflow.test.ts` already exercises full plan/commit workflows on fixture workbooks.
- `packages/sdk/src/index.ts` is the public SDK export surface for workflow helpers.

Folded in a scoped production implementation:

- Added `ReleaseProofBundle` and related types.
- Added `createReleaseProofBundle(plan, commit, options)` in `packages/sdk/src/agent-workflow.ts`.
- Exported the helper and types from `packages/sdk/src/index.ts`.
- Added a fixture-backed test that creates a workbook, plans an edit, commits with `expectSha256`, reopens the output, diffs before/after, builds a release proof bundle, and verifies consistency checks.

The helper is deliberately conservative:

- It uses existing `AgentPlanResult` and `AgentCommitResult` evidence.
- It accepts optional diff evidence instead of reopening files itself.
- It records local-proof claim boundaries and does not imply signed provenance.
- It reports failed consistency checks instead of throwing, so blocked/degraded evidence can still be surfaced.

## Results

Validation passed:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "release proof bundle"
bun test packages/sdk/src/agent-workflow.test.ts
bunx biome check packages/sdk/src/agent-workflow.ts packages/sdk/src/agent-workflow.test.ts packages/sdk/src/index.ts
bunx tsc --build
bun run test:changed
```

`bun run test:changed` expanded to the full suite and passed:

| Metric | Count |
| --- | ---: |
| Passing tests | 4904 |
| Skipped tests | 1 |
| Failing tests | 0 |
| Assertions | 27625 |

The new helper creates a bundle with:

- subject: file, output, input SHA-256, output SHA-256, plan digest;
- operations: count and plan/commit operation artifact digests;
- plan: trace digest, phases, artifact digests, check/lint/audit/write-policy booleans;
- commit: trace digest, output hash, phases, artifact digests, check/lint/audit/write-policy booleans;
- reopen: post-write validity, reopen status, output hash, check/lint/package graph audit status;
- diff: optional sheet diff evidence;
- consistency: named checks and issues;
- claim boundaries: local evidence, not signed provenance.

## Confidence

High. This is a small wrapper around already-tested workflow evidence, and the validation suite passed. The remaining uncertainty is schema stability for external consumers, not implementation correctness.

## Fold-in decision

Folded into production SDK.

Next fold-in should add compact CLI/API/MCP exposure, but only after deciding whether release proof bundles are returned by default, behind a flag, or written as explicit artifacts.

## Next question

Can formula corpus mismatch records be folded into production with explicit oracle route fields and counters?
