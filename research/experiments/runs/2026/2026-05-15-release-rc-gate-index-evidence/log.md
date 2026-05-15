# Release RC Gate Index Evidence

## Question

Should `release-proof-index` consume the unified `release:rc:gate` command as release packageability evidence now?

## Hypothesis

Yes, if the index records it as local packageability evidence only and keeps publication, provenance, lifecycle, protocol-session, retention, and privacy policy gates missing.

## External sources checked

- npm package provenance documentation: https://docs.npmjs.com/generating-provenance-statements
- npm provenance viewing documentation: https://docs.npmjs.com/viewing-package-provenance/
- GitHub artifact attestation documentation: https://docs.github.com/en/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds
- Bun install documentation: https://bun.sh/docs/cli/install

## Why this matters to Ascend

The release proof bundle should make packageability evidence discoverable from the same owner-handoff artifact as safe-open and package-action proof. It should not imply registry publication, signed provenance, SLSA, GitHub attestations, production API readiness, or full MCP protocol compatibility.

## Probe/implementation

Added `rcGateCommand: "bun run release:rc:gate"` to `releasePackageabilityEvidence` in the release proof index and owner-handoff JSON. The covered evidence now names the unified RC gate's build, pack, isolated install, dependency-leak rejection, and SDK/CLI/API/MCP workbook proof.

## Results

Targeted validation:

- `bun test fixtures/benchmarks/release-proof-index.test.ts`
- `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json`
- `bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts research/experiments/index.md research/experiments/runs/2026/2026-05-15-release-rc-gate-index-evidence/log.md`
- `bunx tsc --build`
- `bun run test:changed`

The proof shape keeps `ownerApprovalRequired=true`, preserves the local-tarball status, and leaves all publication/provenance/policy blockers intact.

## Confidence

Medium-high for owner routing. The evidence points to a passing local RC gate but does not embed or attest the RC output.

## Fold-in decision

Promote to release loop as proof-index evidence. Do not promote to headline release wording.

## Next question

Can the release proof index expose a minimal publication-policy checklist for npm provenance versus explicit non-provenance wording without adding fake attestation claims?
