# Release Owner Action Markdown Table

## Question

Should the release proof Markdown render owner acceptance evidence as a readable table, or is the compact `--owner-handoffs-json` output sufficient?

## Hypothesis

The Markdown proof report should render the owner actions as a table. Owner loops often review Markdown first, while automation can consume JSON. A table makes the same acceptance evidence and forbidden shortcuts visible without adding any SDK, CLI, API, MCP, or compact top-claim report surface.

## External sources checked

- SLSA verifying artifacts: https://slsa.dev/spec/v1.0-rc1/verifying-artifacts
- GitHub CLI `gh attestation verify`: https://cli.github.com/manual/gh_attestation_verify
- Sigstore bundle format: https://docs.sigstore.dev/about/bundle/
- Sigstore signature verification: https://docs.sigstore.dev/cosign/verifying/verify/

## Why this matters to Ascend

The release claim board is now intentionally conservative. The next risk is not missing proof; it is a human reviewer misreading a digest or generated fixture as stronger evidence than it is. A Markdown owner-action table keeps the proof legible and keeps the forbidden shortcuts adjacent to each release gate.

## Probe/implementation

Added a `## Next Owner Actions` table to `releaseProofIndexMarkdown` with:

- rank;
- artifact;
- gate;
- owner loop;
- priority;
- next step;
- acceptance evidence;
- forbidden shortcut.

This keeps the existing one-line readiness summary for backwards-compatible scanning and adds a human-readable table below it. Compact safe-open and package-action reports were not changed.

## Results

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings
```

The release-index tests now assert the table header plus representative product and release rows. The generated Markdown shows product fixture acceptance, performance latency evidence requirements, correctness unsupported-feature boundaries, release provenance/publication boundaries, and compact-report publication policy in table form.

## Confidence

High that the table improves owner readability without changing proof semantics. It is still local proof, not a signed release artifact or product surface.

## Fold-in decision

Promote to proof harness/reporting. Keep top handoffs unchanged: safe unknown workbook opening and auditable package-part mutation only. Keep formula rename frozen.

## Next question

Can the owner-action table be used as the single human handoff in the next product/performance release-proof loop, or does that loop still need separate release notes copy?
