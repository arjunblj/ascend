# Safe Open Release Index Blockers

## Question

Can the safe unknown workbook opening proof bundle become more release-ready without adding another product surface?

## Hypothesis

Yes. The release proof index already references safe-open by digest, but it should also carry the exact reproduction command and the honest publication blockers so release work does not overclaim local proof as public certification.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Trusted Documents: https://support.microsoft.com/en-us/office/trusted-documents-cf872bd8-47ec-4c02-baa5-1fdba1a11b53
- Microsoft Safe Documents: https://support.microsoft.com/en-us/office/safe-documents-e2071599-fb31-442b-a30c-198c25e2aacd
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

Safe unknown workbook opening is the top claim handoff. The claim is strong only if release evidence remains reproducible and bounded: Ascend routes package features before hydration; it does not provide malware scanning, sandboxing, file trust, signed provenance, or release performance thresholds.

## Probe/implementation

Inspected the current safe-open proof harness, release proof index, and safe-open synthesis. Reran the proof:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1
bun run fixtures/benchmarks/release-proof-index.ts --no-timings
```

Folded in a small benchmark-harness improvement:

- `fixtures/benchmarks/release-proof-index.ts` now includes each artifact's reproduction command.
- It adds `publicationStatus` and `publicationBlockers`.
- The safe-open artifact remains `needs-release-packaging` because signed and unknown-part cases are code-generated packages, not public binary fixtures, and local timings are proof-run observations rather than thresholds.

No SDK, CLI, API, MCP, or production open-plan surface was changed.

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts fixtures/benchmarks/safe-open-proof.test.ts packages/sdk/src/open-plan.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts research/experiments/syntheses/2026-05-safe-open-proof-bundle.md research/experiments/index.md research/experiments/runs/2026/2026-05-15-safe-open-release-index-blockers/log.md
bunx tsc --build
bun run test:changed
```

`bun run test:changed` detected a repository-level change and ran the full test suite: 5048 passed, 1 skipped, 0 failed.

## Results

Safe-open proof rerun:

- 9 cases.
- 8 OK.
- 1 malformed package rejected.
- 4 review-before-hydration cases.
- Public fixture open-plan/full-open ratios: 14.09x to 31.97x.

Release proof index still contains only the top two artifacts:

- `safe-open-proof`
- `package-action-proof`

The index now prints commands such as:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json
```

and explicit publication blockers instead of letting release notes infer readiness from the digest alone.

Latest no-timings stable shape digests:

- `safe-open-proof`: `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`
- `package-action-proof`: `b9758496346c97920c80ba08b6632315708a6d6cc770927695337e729554dbb0`

## Confidence

High for the fold-in: it is testable metadata in the proof harness and does not change workbook behavior. Medium for external publication readiness because public signed/unknown binary fixture replacement is still unresolved.

## Fold-in decision

Promote to product/performance proof packaging. Keep production open-plan unchanged. Do not publish safe-open as a stronger claim until the blockers are accepted or resolved.

## Next question

Can the auditable package-part mutation artifact receive the same release-readiness treatment while preserving the no-attestation and chart-byte-boundary language?
