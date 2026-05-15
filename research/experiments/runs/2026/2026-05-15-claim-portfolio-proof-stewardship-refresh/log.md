# Claim Portfolio Proof Stewardship Refresh

## Question

After the recent tiny fold-ins, does the ranked research portfolio still point to the same top one or two highest-leverage unknowns, and can we prove that without promoting another surface?

## Hypothesis

Yes. Current proof artifacts should show that safe unknown workbook opening and auditable package-part mutation remain the only implementation-loop handoffs; formula intelligence stays rejection-first, and practical latency remains diagnostic-only.

## External sources checked

- LSP 3.17 separates `prepareRename` from edit-producing rename, matching Ascend's refusal-first formula boundary: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft names in formulas document workbook- and worksheet-scoped names, which require workbook context before rename can be safe: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft structured references document table and column symbols in formulas: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Bun profiling docs support keeping phase profiles as diagnostic performance evidence until a release environment approves thresholds: https://bun.sh/docs/project/benchmarking

## Why this matters to Ascend

The highest-value research output right now is not another narrow surface. It is deciding what Ascend can credibly claim, what proof blocks stronger wording, and which owner loop should take the next step. This keeps research from competing with correctness, performance, or product loops.

## Probe/implementation

No product surface was added. I reran existing proof artifacts and inspected the current synthesis:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json
bun run fixtures/benchmarks/practical-latency-contracts.ts --dry-run --contract edit-verify --repeat 1 --warmup 0 --json --out-dir /tmp/ascend-phase-profile-dry-run
```

## Results

Release proof index:

- `artifactCount`: 2
- `excludedEvidenceCount`: 1
- `releaseGate`: `blocked-by-publication-policy`
- `headlineClaimsAllowed`: `false`
- `missingRequirementCount`: 9
- missing owner loops: correctness 1, performance 2, product 2, release 4

Safe-open compact proof:

- 9 cases: 6 public fixtures, 2 synthetic edge packages, 1 malformed package
- 8 OK, 1 rejected
- 4 review-before-hydration cases
- malformed rejection: true
- headline claim allowed: false

Package-action compact proof:

- 8 cases: 2 public fixtures, 2 generated workbooks, 4 generated edge packages
- action counts: passthrough 27, regenerate 38, add 3, drop 3, error 1
- all action classes covered: true
- source graph everywhere: true
- streaming proof cases: 1
- headline claim allowed: false

Formula-assist proof:

- public formulas discovered/sampled: 1685/1685
- reference spans: 2322
- binding roles: 25
- LET-local prepare-rename OK targets: 3
- prepare-rename refusals: 1692 by reason
- boundary: no workbook edits, no workbook-context name resolution, no safe cross-workbook/table rename

Practical latency dry-run:

- phase-profile step is wired into edit-verify;
- current run is diagnostic-only due local/private inputs and unrelated tracked engine edits;
- practical latency remains excluded from release proof artifacts.

## Confidence

High that the top-two handoff remains unchanged. Medium that the release gate count is final, because owner loops may approve generated fixtures or boundary language without code changes.

## Fold-in decision

Promote only to topic synthesis and owner-loop handoff. Do not implement formula rename. Do not promote practical latency into release claims. Do not add new SDK/CLI/API/MCP surfaces from this stewardship block.

## Next question

Can the next block stay in claim-steward mode, accepting only top-two owner proof work and tiny correctness fixes that are already in flight?
