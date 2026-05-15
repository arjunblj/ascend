# Claim Board Current Proof Drift Cleanup

## Question

Does the human claim board still match the current machine proof for the top two release claims?

## Hypothesis

Mostly, but stale package-action wording can drift after proof harness improvements. Claim stewardship should update active syntheses from `release-proof-index` rather than keep adding surfaces.

## External sources checked

- SLSA distributing provenance: https://slsa.dev/spec/v1.2/distributing-provenance
- GitHub artifact attestations provenance docs: https://docs.github.com/en/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds
- OpenSSF Scorecard repository and Binary-Artifacts check reference: https://github.com/ossf/scorecard

## Why this matters to Ascend

The release claim board is supposed to prevent overclaiming. If the Markdown says one or two streaming proofs while the machine proof says three, or if stale generated-fixture wording still names public fixture cases, owner loops get unclear instructions. The North Star here is trustworthy mutation planning and honest release wording, not another product surface.

## Probe/implementation

- Reran `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`.
- Searched active syntheses for stale package-action proof text:
  - `1 representative streaming proof case`
  - `Representative streaming proof cases | 1`
  - `two representative streaming proofs`
  - old stable shape `b4cf0755f2fd40ff577ab87b6600ee14a58bbd69608796c903759e7e8b4d29e8`
  - stale product checklist wording that treated `docprops-passthrough` and `calc-chain-drop` as generated edge packages.
- Updated only active synthesis files:
  - `research/experiments/syntheses/2026-05-release-claim-board.md`
  - `research/experiments/syntheses/2026-05-ranked-research-portfolio.md`
  - `research/experiments/syntheses/2026-05-owner-handoff.md`

## Results

- Current release index remains fail-closed: `releaseGate=blocked-by-publication-policy`, `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, `missingRequirementCount=9`.
- Package-action proof current facts:
  - 8 cases;
  - 4 public fixtures, 2 generated workbooks, 2 generated edge packages;
  - action totals `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1`;
  - 3 representative streaming proof cases covering `passthrough`, `regenerate`, `add`, and `drop`;
  - stable compact shape `921dc41d423632c935741dd4fb4e7def4e1c2719c90263c13309e64093419803`.
- Active synthesis stale-text scan now returns no matches for the outdated package-action proof phrases.

## Confidence

High for the active synthesis cleanup. Historical run logs still preserve what was true when they were written, so they were not rewritten.

## Fold-in decision

Promote to topic synthesis only. This is a claim-steward cleanup, not production implementation and not permission to publish stronger release language. Owner gates remain missing.

## Next question

Can the next safe-open or package-action owner loop resolve an actual missing gate, rather than adding more proof-board restatements?
