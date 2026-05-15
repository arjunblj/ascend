# Package Action Journal Compatibility

## Question

Can the auditable package-part proof report correlate package action evidence with rollback-journal package-state issues without adding a new public surface?

## Hypothesis

Yes. The existing package-action proof harness already creates commit results with mutation journals. A report-only compatibility check can show that package action evidence and journal exactness evidence are both present for each proof case.

## External sources checked

- Microsoft Open Packaging Conventions fundamentals define package parts and relationships as the unit of package accounting: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft `System.IO.Packaging` exposes package parts, relationships, and package digital signatures as package-level concepts: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging?view=windowsdesktop-10.0
- SheetJS write options describe writer scope, data preservation orientation, and unsupported-feature boundaries: https://docs.sheetjs.com/docs/api/write-options/
- in-toto attestations bind signed statements to subjects and predicates; Ascend's proof remains local evidence, not an attestation: https://github.com/in-toto/attestation/blob/main/spec/README.md

## Why this matters to Ascend

"Auditable package-part mutation" is credible only when the two safety stories agree: package actions explain emitted parts, and rollback journals explain whether public inverse operations can restore saved package state. This check makes that connection explicit without adding product surface area.

## Probe/implementation

- Inspected `fixtures/benchmarks/package-action-proof.ts` and `packages/sdk/src/agent-workflow.ts`.
- Added report-only fields to each package proof case:
  - `commitJournalExact`
  - `commitJournalPackageIssueCount`
  - `commitJournalPackageIssueRefs`
- Extended the proof Markdown with a `Journal package issues` column.
- Added harness assertions that every release-claim package-action case has non-exact commit journal evidence and at least one `package-part-preservation` issue with refs.
- Did not change SDK, CLI, API, MCP, or writer behavior.

## Results

`bun run fixtures/benchmarks/package-action-proof.ts --no-timings` produced:

- Combined commit actions: `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1`.
- Every case reported `Journal package issues = 1`.
- Only `unknown-part-error` reported a package-action proof issue.

Validation:

```bash
bun test fixtures/benchmarks/package-action-proof.test.ts
```

## Confidence

High for the claim that the existing proof harness now correlates package actions with rollback-journal package-state issues. Medium for release copy until product decides how to publish the report and whether synthetic edge cases need durable public binary fixtures.

## Fold-in decision

Promote to correctness/product proof packaging. This is a tiny proof-harness fold-in, not a new user surface.

## Next question

Can the package-action report include a stable compact JSON fixture for one case without leaking workbook bytes or bloating the repo?
