# Owner Blocker Grouping Proof

## Question

Can Ascend's release claim board make the 9 missing top-claim requirements actionable by grouping them into owner decisions, validation runs, optional harness expansion, and publication policy, without adding another product surface?

## Hypothesis

Yes. `release-proof-index --json` already emits enough `nextOwnerActions` data to group every blocker by owner loop, artifact, priority, and `nextStepKind`. The synthesis should point owners at those existing fields instead of inventing a new workflow.

## External sources checked

- Microsoft Protected View documentation anchors safe-open contrast and trust-boundary wording: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions documentation anchors package-action part/relationship/signature boundaries: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- SheetJS write options document preservation/write-scope boundaries for competitor contrast: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

The current blockers are not invitations for research to create more SDK, CLI, API, or MCP surfaces. They are mostly owner approvals and publication policy. A blocker closure map helps the next loops close claims by making the remaining proof operational.

## Probe/implementation

Ran:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

Then folded the owner grouping into the release claim board and experiment index. No production code changed.

## Results

The release index reports:

- `missingRequirementCount=9`
- `missingByOwnerLoop`: correctness 1, performance 2, product 2, release 4
- `missingByArtifact`:
  - safe-open: `public-edge-fixtures`, `release-latency-run`, `publication-boundary`, `compact-report-publication-policy`
  - package-action: `edge-fixture-policy`, `provenance-boundary`, `unsupported-feature-boundary`, `streaming-matrix-boundary`, `compact-report-publication-policy`
- `nextStepKind` groups:
  - owner decision or fixture replacement: 2 blockers
  - owner boundary approval: 1 blocker
  - validation run: 1 blocker
  - owner decision or harness expansion: 1 blocker
  - publication policy: 4 blockers

This means 7 of 9 blockers are owner decisions or publication policy, 1 is a performance validation run, and 1 is a performance choice between accepting narrow wording or expanding harness coverage.

## Confidence

High. The proof is generated directly from the canonical release index and does not depend on private workbooks or local timing thresholds.

## Fold-in decision

Promote to topic synthesis only. Do not add a new gate or product surface. Future owner loops should use the existing `nextOwnerActions`, `implementationHandoffs`, and `readyWhen` fields.

## Next question

Should the safe-open owner resolve the `public-edge-fixtures` gate by accepting disclosed generated structural packages, or keep searching for public signed/unknown binary fixtures before any headline release wording?
