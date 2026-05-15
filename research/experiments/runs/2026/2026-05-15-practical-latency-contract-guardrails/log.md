# Practical Latency Contract Guardrails

## Question

Can Ascend's practical latency contract report stop benchmark evidence from being mistaken for a release claim when the worktree is dirty or when edit-verify time is hidden in unassigned prepared-commit overhead?

## Hypothesis

Yes. The benchmark report can remain a performance/product proof artifact if it records worktree state, labels dirty reports as diagnostic, and attributes prepared commit time to named phases plus an explicit unassigned/finalize bucket.

## External sources checked

- [Hyperfine README](https://github.com/sharkdp/hyperfine) documents warmup runs and repeated measurements, reinforcing that benchmark output needs explicit run context before comparison.
- [Bun benchmarking docs](https://bun.com/docs/project/benchmarking) document CPU and heap profiling outputs, supporting the existing profile-command handoff in the report.
- [SLSA attestation model](https://slsa.dev/spec/v1.1/attestation-model) frames evidence as metadata about an artifact; this is the boundary that benchmark reports should not cross into signed provenance claims.

## Why this matters to Ascend

The ranked portfolio currently treats practical latency as supporting evidence for safe unknown workbook opening and trustworthy agent mutation. If latency reports are generated from an undocumented dirty worktree or hide a large prepared-commit remainder, they can create false confidence and weak handoffs.

## Probe/implementation

Inspected the committed benchmark fold-in `b495af20 perf(benchmarks): guard latency contract evidence`, which changed `fixtures/benchmarks/practical-latency-contracts.ts` to:

- include `git status --short --branch` state in JSON and Markdown summaries,
- mark tracked-dirty runs as "not release-claimable",
- record untracked entry counts separately from tracked code changes,
- include a worktree guardrail section in both summary reports,
- add "Prepared commit unassigned/finalize overhead" to the edit-verify decision matrix.

Ran a local dry-run probe:

```bash
bun run fixtures/benchmarks/practical-latency-contracts.ts --dry-run --json --contract edit-verify --out-dir /private/tmp/ascend-contract-dry-run
```

The probe emitted `worktree.trackedDirty: false`, `untrackedCount: 20`, summary artifact paths, and skipped edit-verify steps with profile commands intact.

## Results

- The earlier full-suite dense-reader failure no longer reproduced:
  - `bun test packages/io-xlsx/src/reader/reader.test.ts -t "keeps parsed narrow dense sheets in dense chunks"` passed.
  - A handcrafted XLSX probe for `A1:E160` also hydrated into dense chunks.
- `bun run test:changed` now passes the full affected suite: 5059 pass, 1 skip, 0 fail.
- The latency-contract guardrail is implemented and committed in `b495af20`.

## Confidence

Medium. The dry-run proves report shape and worktree metadata without running expensive benchmark payloads. Real latency numbers still require a non-dry-run contract run on a clean tracked worktree.

## Fold-in decision

Promote to performance/product loop as benchmark evidence hygiene. Do not promote to a product speed claim by itself.

Allowed wording: "practical latency reports now record worktree cleanliness and surface prepared-commit unassigned overhead."

## Next question

Should the release proof index consume latency-contract summaries only when `trackedDirty === false`, or should latency reports remain outside release proof until public corpus inputs are standardized?
