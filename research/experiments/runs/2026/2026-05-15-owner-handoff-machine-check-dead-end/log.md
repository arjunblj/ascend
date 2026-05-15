# Owner Handoff Machine Check Dead End

## Question

Can the owner handoff be machine-checked by making a small Markdown-only acceptance table in the release claim board the single source of truth, or would that duplicate `release-proof-index`?

## Hypothesis

It would duplicate `release-proof-index`. Markdown checkboxes are useful for human owner approval, but the release proof index already owns machine-readable gates, owner loops, priorities, and artifact mapping.

## External sources checked

- Write the Docs docs-as-code guide: https://www.writethedocs.org/guide/docs-as-code.html
- Diataxis documentation framework: https://diataxis.fr/
- Google developer documentation style guide for code in text: https://developers.google.com/style/code-in-text
- Google Developer Knowledge API announcement: https://developers.googleblog.com/introducing-the-developer-knowledge-api-and-mcp-server/

## Why this matters to Ascend

Claim stewardship needs durable gates, but adding a second checked Markdown table would create drift risk. The product needs one machine source of truth and one human handoff, not two competing gate models.

## Probe/implementation

- Ran a local Bun probe over `runReleaseProofIndex`.
- Parsed `research/experiments/syntheses/2026-05-owner-handoff.md` for unchecked owner checklist items.
- Compared canonical `readyWhen` gate count against human checkbox count.
- Updated the owner handoff with an explicit source-of-truth rule.

## Results

| Field | Value |
| --- | --- |
| Canonical `readyWhen` gates | 9 |
| Owner handoff checkboxes | 11 |
| Machine source of truth | `fixtures/benchmarks/release-proof-index.ts` |
| Human checklist | `research/experiments/syntheses/2026-05-owner-handoff.md` |

Canonical gates:

- `safe-open-proof/public-edge-fixtures`
- `safe-open-proof/release-latency-run`
- `safe-open-proof/publication-boundary`
- `safe-open-proof/compact-report-publication-policy`
- `package-action-proof/edge-fixture-policy`
- `package-action-proof/provenance-boundary`
- `package-action-proof/unsupported-feature-boundary`
- `package-action-proof/streaming-matrix-boundary`
- `package-action-proof/compact-report-publication-policy`

Decision: do not add a Markdown-only machine gate. The 11 owner checkboxes deliberately split decisions more finely than the 9 release-index gates.

## Confidence

High. A second machine-readable Markdown acceptance table would increase maintenance cost without improving proof quality.

## Fold-in decision

Archive as dead end. Keep `release-proof-index.ts` as canonical machine gate and `2026-05-owner-handoff.md` as the human approval checklist.

## Next question

Can we now stop adding stewardship structure and rerun the top proof commands once more to produce a concise status checkpoint for the next owner loops?
