# Release Handoff Proof Requirements

## Question

Can the top-two release proof handoffs carry the full claim ladder as data: fixture, benchmark, surface, validation gate, competitor contrast, honest boundary, and kill criterion?

## Hypothesis

Yes. The release proof index is already the canonical machine-readable artifact for top claim readiness. Adding proof-requirement fields to its implementation handoffs makes the next owner loops start from proof and refusal conditions instead of broad research prose.

## External sources checked

- Microsoft Protected View documents read-only handling for potentially unsafe files; this is the competitor contrast for safe-open trust UX versus Ascend package-feature routing: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions describe packages as parts and relationships, including digital signatures; this frames Ascend's per-part proof boundary: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl documents that unsupported objects such as shapes can be lost when opening and saving, which supports contrasting preservation-boundary docs against Ascend per-part accounting: https://openpyxl.pages.heptapod.net/openpyxl/tutorial.html
- SheetJS CE write options document writer scope and data-preservation orientation, which supports the package-action competitor contrast: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

The user asked for product-shaped claims with proof requirements and kill criteria, not feature-shaped surfaces. The previous release proof index named the top two handoffs and blockers, but the proof ladder still lived mostly in synthesis prose. Encoding the proof ladder in the handoff data makes the implementation loops auditable and keeps lower-ranked surfaces frozen.

## Probe/implementation

Added `ReleaseProofClaimProofRequired` to `fixtures/benchmarks/release-proof-index.ts` and attached it to each `ReleaseProofImplementationHandoff`.

Each top handoff now includes:

- fixture requirement;
- benchmark requirement;
- surface boundary;
- validation gate;
- competitor contrast;
- honest boundary;
- kill criterion.

The Markdown handoff line now includes the kill criterion, and tests assert both safe-open and package-action proof ladders.

Commands run:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

## Results

- Focused test passed: 3 tests, 73 assertions.
- JSON output now includes `proofRequired` for `safe-open-proof` and `package-action-proof`.
- Both handoffs still report `implementationSurfacePromotionAllowed=false`.
- No new SDK, CLI, API, MCP, package, or app surface was added.

## Confidence

High for the handoff shape and top-two scope. Medium for final wording until product/release owners approve the exact claim copy.

## Fold-in decision

Fold into the release proof harness. This is proof routing and claim stewardship only.

## Next question

Can the same release proof index emit a compact owner-ready table for acceptance checkboxes without duplicating the full synthesis document?
