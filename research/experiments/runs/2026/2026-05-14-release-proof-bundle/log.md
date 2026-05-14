# Release Proof Bundle

## Question

Can we build a release proof bundle that demonstrates inspect, plan, commit, reopen, diff, and audit on real files without fake claims?

## Hypothesis

Yes. Ascend already emits trace digests, artifact digests, package graph audits, write policy diagnostics, post-write reopen verification, output hashes, check/lint results, and compact summaries. A release proof bundle can organize those existing facts into a stable artifact, as long as it clearly distinguishes local evidence from signed supply-chain provenance.

## External sources checked

- [SLSA Provenance](https://slsa.dev/spec/v1.0-rc1/provenance): provenance records build definitions, external parameters, resolved dependencies, run details, and artifact digests.
- [SLSA software attestations](https://slsa.dev/spec/v1.0/attestation-model): attestations make metadata explicit instead of relying on a raw signature to imply meaning.
- [GitHub artifact attestations](https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations): GitHub Actions can generate attestations for build artifacts and SBOMs, then verify them with the GitHub CLI.
- [OpenSSF Scorecard](https://github.com/ossf/scorecard): automated checks can help users assess OSS project security posture, but the project explicitly treats scores as heuristics with false positives and false negatives.

## Why this matters to Ascend

Ascend's strongest positioning is trustworthy workbook mutation: inspect what exists, plan an edit, explain preservation/loss, commit, reopen, diff, and audit the output. Current OSS spreadsheet tools often expose APIs but do not package those claims as evidence. A release proof bundle could become Ascend's differentiator for agents and humans: every release or demo file can carry a machine-readable chain of facts instead of a marketing claim.

## Probe/implementation

Inspected local implementation:

- `packages/sdk/src/agent-workflow.ts` defines `AgentWorkflowTrace`, `AgentTraceArtifact`, `traceDigest`, `artifact` digests, `digestPlan`, `createAgentPlan`, `commitAgentPlan`, `compactAgentPlanResult`, and `compactAgentCommitResult`.
- Planning already records preview, check, lint, loss audit, package graph audit, approval audit, preservation summary, and write policy diagnostics.
- Commit already records input/output hashes, hash guard, apply, recalc, write, post-write reopen, post-write package graph roundtrip audit, check, lint, and compact post-write summaries.
- `verifyWrittenWorkbook` reopens the output workbook and audits package graph roundtrip preservation.
- CLI/API tests already assert compact commit evidence such as post-write validity, output hash agreement, check validity, package graph audit status, and trace artifact counts.

Added ignored probe `research/experiments/runs/2026/2026-05-14-release-proof-bundle/probes/release-proof-candidate.ts`. It:

1. Uses public fixture `fixtures/xlsx/poi/SampleSS.xlsx`.
2. Inspects the input workbook.
3. Creates an agent plan for one `setCells` operation.
4. Commits the plan to an ignored output workbook with `expectSha256`.
5. Reopens the output workbook.
6. Diffs input vs output.
7. Writes an ignored proof JSON bundle with input, plan, commit, reopen, diff, audit, digest, and claim-boundary sections.
8. Fails if hash guards, plan digest linkage, post-write reopen, post-write audits, package graph audits, or diff evidence are missing.

Validation command:

```bash
bun run research/experiments/runs/2026/2026-05-14-release-proof-bundle/probes/release-proof-candidate.ts
```

## Results

The probe passed and produced ignored artifacts under the run's `probes/artifacts` directory.

| Evidence | Value |
| --- | --- |
| Input workbook | `fixtures/xlsx/poi/SampleSS.xlsx` |
| Input SHA-256 | `44f1b6ef310c370d4902ee3452e6174da25a67bdb09f47850d51b3287cb3db71` |
| Output SHA-256 | `4aba57595e558b911bb518d66f73fc94df5cd91bc0c436382a09385648fd8ac1` |
| Plan digest | `6e2ee0e82b62cc79ebc7ba19eba2f79cdf9c1dec0e8a7c94229a0eee9327f1f7` |
| Plan trace digest | `8e3182f1ce2b2c1dc1fce293f3eb933a4942873725b14324082ce65eaa1e3113` |
| Commit trace digest | `f5385364c387ef51f7710d7afd0e3af960770ab78654b1eb21aa89296b0b324e` |
| Post-write audits passed | `true` |
| Diff sheets | `1` |

The generated proof JSON included claim boundaries:

- local proof bundle, not a signed SLSA attestation;
- public repo fixture only, no private corpus workbook copied;
- output workbook and proof JSON are ignored probe artifacts;
- bundle reports actual plan, commit, reopen, diff, and audit results from this run.

This is already close to a production-worthy proof bundle. The main gap is schema hardening, artifact storage policy, optional signing/attestation, and keeping bundle size bounded for CLI/API/MCP output.

## Confidence

High that Ascend has enough internal evidence to build a release proof bundle. Medium that the exact schema is right; it should be treated as a candidate and aligned with SLSA/in-toto terminology before external claims.

## Fold-in decision

Promote to product/DX loop and correctness loop.

Recommended fold-in:

1. Add a `createReleaseProofBundle` SDK helper that takes existing `AgentPlanResult` and `AgentCommitResult`.
2. Include stable sections: subject, input, operations, plan, commit, reopen, diff, audits, artifacts, limitations.
3. Add a compact mode for CLI/API/MCP and a full JSON mode for release artifacts.
4. Keep claims explicit: "local evidence" by default; "signed provenance" only when a real attestation exists.
5. Add tests on a public fixture workbook and on an unsupported-feature workbook where the bundle must report blocked or degraded evidence rather than passing.
6. Later, map the bundle to SLSA provenance or GitHub artifact attestations for release builds.

Do not fold into production in this research loop. The next product/correctness loop should define the schema and wire tests.

## Next question

Which fold-in candidate should be promoted first across correctness, performance, and product/DX loops?
