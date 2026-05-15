# Safe Open Compact Digest Policy

Date: 2026-05-15

## Question

Should safe-open compact reports stay generated on demand until the same artifact storage, privacy, and canonicalization policy exists for package-action compact reports?

## Hypothesis

Yes. The safe-open compact report is useful as a release handoff artifact, but its raw JSON digest should not be indexed until release owners define the subject bytes, storage location, stable canonicalization rule, and wording boundaries.

## External sources checked

- GitHub artifact attestations require a subject digest for the attested artifact: https://docs.github.com/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds
- SLSA provenance includes artifact digests and verification expectations: https://slsa.dev/spec/v1.0-rc1/provenance
- SLSA artifact verification discusses matching provenance to artifacts: https://slsa.dev/spec/v1.0-rc1/verifying-artifacts
- Microsoft OPC digital signatures validate signed package contents: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/digital-signatures

## Why this matters to Ascend

Safe-open is the top product/performance claim. Adding digest links too early could make a generated report look like retained artifact evidence or attestation. The release proof index should stay honest until storage and publication semantics are explicit.

## Probe/implementation

I generated the safe-open compact report twice:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json
```

Local probe results:

| Check | Result |
| --- | --- |
| Compact JSON bytes | 4959 each run |
| Raw SHA-256 | different across runs |
| Stable SHA-256 after deleting `generatedAt` | `0f614b88ca4c5ea0f2a8c4c54147682be144dc404d12502840e770e21d036ae8` both runs |

No production code was changed.

## Results

The compact safe-open report is structurally stable but not byte-stable as generated output. This matches the package-action compact report policy: useful generated proof, not yet a release-index artifact by digest.

## Confidence

High. The local digest probe is direct, and external attestation systems treat digest publication as a subject-artifact commitment.

## Fold-in decision

Do not fold safe-open compact report digests into the release proof index yet. Keep compact reports generated on demand until release owners define artifact storage, privacy filtering, and canonicalization for all top-claim compact reports.

## Next question

Should the claim board rank compact report publication policy as a shared release-loop blocker for both top claims?
