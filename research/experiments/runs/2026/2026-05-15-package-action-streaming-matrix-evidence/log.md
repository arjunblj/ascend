# Package Action Streaming Matrix Evidence

## Question

Can the auditable package-part mutation claim expose exactly what its single streaming writer proof covers, and what remains unproven, in machine-readable release handoff data?

## Hypothesis

Yes. The existing package-action proof already records streaming action counts for the representative `docprops-passthrough` case. The release proof index can derive a fail-closed `streamingMatrixEvidence` object that lists covered action kinds, missing action kinds, covered cases, non-streaming cases, and public non-streaming cases without satisfying the performance owner gate.

## External sources checked

- Node.js streams documentation: https://nodejs.org/api/stream.html
- Bun streams documentation: https://bun.sh/docs/api/streams
- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

The North Star claim is auditable package-part mutation: every write should be explainable as local package evidence without overclaiming preservation or performance. Streaming is attractive for large workbooks, but one representative dirty-sheet streaming proof is not the same as full parity across `passthrough`, `regenerate`, `add`, `drop`, and `error` cases. Owner loops need the boundary as data, not narrative.

## Probe/implementation

- Inspected `fixtures/benchmarks/package-action-proof.ts` and current release-proof index code.
- Ran:

```bash
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json | jq '{cases:[.cases[] | {name, sourceKind, streaming: (.streamingProof != null), streamingActionCounts: .streamingProof.actionCounts}], combinedCommitActionCounts}'
```

- Added `streamingMatrixEvidence` to `fixtures/benchmarks/release-proof-index.ts`, the owner-handoff JSON, and the release proof Markdown.
- Added regression assertions in `fixtures/benchmarks/release-proof-index.test.ts`.
- Updated the ranked portfolio, release claim board, owner handoff, and experiment index.

## Results

Current proof data:

| Field | Value |
| --- | --- |
| Streaming proof cases | `1` |
| Covered case | `docprops-passthrough` |
| Covered action kinds | `passthrough`, `regenerate` |
| Missing action kinds | `add`, `drop`, `error` |
| Non-streaming cases | `regenerate-existing-sheet`, `add-sheet-part`, `calc-chain-drop`, `signature-invalidation-drop`, `macro-passthrough`, `chart-sidecar-accounting`, `unknown-part-error` |
| Public non-streaming cases | `calc-chain-drop`, `macro-passthrough`, `chart-sidecar-accounting` |
| Owner gate | `streaming-matrix-boundary` remains missing |

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
```

Result: 4 tests passed.

## Confidence

High for the release-handoff evidence shape: it is derived directly from the current package-action proof result and covered by tests. Medium for release wording utility: performance still needs to decide whether one representative case is sufficient or whether the streaming matrix must expand.

## Fold-in decision

Promote to performance loop and correctness loop as proof packaging only. This is a tiny release-proof fold-in, not a new SDK, CLI, API, MCP, or writer surface. Keep `streaming-matrix-boundary` missing until performance accepts narrow wording or expands the harness.

## Next question

Should the performance owner accept one representative dirty-sheet streaming proof for narrow wording, or require streaming variants for `add`, `drop`, `error`, public macro, and public chart cases before any release mention?
