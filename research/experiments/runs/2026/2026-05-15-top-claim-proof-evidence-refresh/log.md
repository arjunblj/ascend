# Top Claim Proof Evidence Refresh

## Question

Can the ranked research portfolio prove the top one or two product-shaped unknowns from existing harnesses without promoting new SDK, CLI, API, MCP, formula-rename, viewport, agent-view, sidecar, oracle, or observability surfaces?

## Hypothesis

Yes. The release proof index and the two compact proof harnesses should provide enough evidence to hand off only safe unknown workbook opening and auditable package-part mutation while keeping every other direction below implementation promotion.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Open Packaging Conventions fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- SheetJS CE write options: https://docs.sheetjs.com/docs/api/write-options/
- Microsoft `System.IO.Packaging` package/signature APIs: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging?view=windowsdesktop-10.0

## Why this matters to Ascend

Ascend's North Star needs release-safe claims, not an expanding research backlog. Safe open and package-action proof are the highest-leverage claims because they connect preservation-first XLSX handling, trustworthy mutation planning, real-world performance evidence, and agent-readable auditability. The external references reinforce the boundaries: Protected View is a trust/security UX, OPC is package graph accounting, and mainstream libraries document writer/preservation limits rather than per-part mutation proof.

## Probe/implementation

Ran the current proof commands:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
```

Updated:

- `research/experiments/syntheses/2026-05-ranked-research-portfolio.md`
- `research/experiments/index.md`

No production code or public product surface changed.

## Results

`release-proof-index --owner-handoffs-json` reported:

- `releaseGate=blocked-by-publication-policy`
- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- exactly two top implementation handoffs: `safe-open-proof` and `package-action-proof`

Safe unknown workbook opening proof:

- 9 cases
- 8 OK, 1 malformed rejection
- 6 public fixtures, 2 generated edge packages, 1 malformed package
- 4 review-before-hydration routes
- risk families: `preservedActiveX`, `preservedMacro`, `preservedOther`, `preservedSignature`
- missing owner gates: `public-edge-fixtures`, `release-latency-run`, `publication-boundary`

Auditable package-part mutation proof:

- 8 cases
- 4 public fixtures, 2 generated workbooks, 2 generated edge packages
- action totals: `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1`
- source graph evidence and package-preservation journal issues in every case
- 1 representative streaming proof case
- missing owner gates: `edge-fixture-policy`, `provenance-boundary`, `unsupported-feature-boundary`, `streaming-matrix-boundary`

The proof supports two handoffs only. It does not support formula rename, new agent-view surfaces, new viewport surfaces, columnar sidecar production work, oracle compatibility claims, or signed provenance language.

## Confidence

High for the ranking and no-promotion decision because it comes from current machine-readable proof outputs and compact harness reruns. Medium for release readiness because nine owner gates remain open, including fixture policy, release latency, unsupported-feature boundaries, streaming wording, and publication/provenance policy.

## Fold-in decision

Promote to topic synthesis and owner handoff only. Do not fold into production. The top two implementation loops should package and validate existing proof, not expand Ascend's public surface area.

## Next question

Can product owners resolve the generated-fixture policy blockers for safe-open and package-action proof without weakening the honest boundary between package-topology evidence and public binary real-world evidence?
