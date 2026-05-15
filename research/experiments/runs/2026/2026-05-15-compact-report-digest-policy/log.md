# Compact Report Digest Policy

Date: 2026-05-15

## Question

Should the release proof index link compact per-claim reports by digest now, or should compact reports stay generated on demand until artifact storage and privacy policy are decided?

## Hypothesis

Compact reports should stay generated on demand. A raw digest over the generated JSON will churn because the report includes `generatedAt`, while a stable-shape digest would be meaningful only after Ascend defines artifact storage, privacy filtering, and publication semantics.

## External sources checked

- GitHub artifact attestations bind a subject digest to a provenance statement: https://docs.github.com/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
- SLSA provenance verification expects the statement subject to match the artifact digest: https://slsa.dev/spec/v1.0-rc1/provenance
- in-toto defines attestation statements and predicate types: https://github.com/in-toto/attestation
- Microsoft OPC digital signatures validate signed package contents: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

Ascend needs release evidence that is honest about what is proven. Digesting compact reports in the release index could look like artifact attestation if the project has not defined where reports live, which bytes are the subject, whether generated fields are stripped, and how private workbook evidence is excluded.

## Probe/implementation

I generated the compact package-action report twice:

```bash
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
```

Local probe results:

| Check | Result |
| --- | --- |
| Compact JSON bytes | 5704 each run |
| Raw SHA-256 | different across runs |
| Stable SHA-256 after deleting `generatedAt` | `55410ea37d67391bd97dd431cfa53fe3fe7ffab7bda8e1f8b53d5f01314b30c7` both runs |

No production code was changed.

## Results

The compact report is deterministic in shape but not byte-stable as generated output. Indexing its raw digest would create churn without adding proof value. Indexing a stable-shape digest would be technically possible, but doing that now would promote an artifact-storage decision that product/release have not made.

## Confidence

High. The local probe directly measured digest behavior, and external provenance systems make clear that digest publication should name the subject artifact and verification expectations.

## Fold-in decision

Do not fold into the release proof index yet. Keep compact reports generated on demand. Promote digest indexing only after a release owner defines artifact storage, privacy filtering, stable canonicalization, and wording that does not imply signed provenance.

## Next question

Can the safe-open proof get the same compact release-report treatment without adding a new open surface or weakening the public-fixture blocker?
