# Package Action Provenance Boundary Audit

## Question

Can the package-action `provenance-boundary` gate become approval-ready without adding mutation surfaces or implying signed provenance?

## Hypothesis

Yes. The current proof artifacts already fail closed: they expose local digest evidence and readiness gates, but they do not claim SLSA, in-toto, GitHub artifact attestation, Sigstore, signed provenance, or third-party verification. The useful work is to make the boundary explicit for the release owner.

## External sources checked

- SLSA specification v1.2: https://slsa.dev/spec/v1.2/
- GitHub artifact attestations for build provenance: https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
- Sigstore overview: https://docs.sigstore.dev/
- Sigstore cosign attest command docs: https://github.com/sigstore/cosign/blob/main/doc/cosign_attest.md
- in-toto Statement layer specification: https://github.com/in-toto/attestation/blob/main/spec/v1/statement.md

## Why this matters to Ascend

Auditable package-part mutation is a top release claim. If local proof digests are described as provenance or attestation, Ascend would overclaim trust properties that belong to a release signing and verification pipeline, not the XLSX package writer.

## Probe/implementation

- Inspected `fixtures/benchmarks/release-proof-index.ts`, `fixtures/benchmarks/package-action-proof.ts`, and the current release claim board.
- Ran `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json`.
- Ran `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json`.
- Ran a local Bun audit over both harness APIs to assert:
  - `signed=false`
  - `attestation=false`
  - `headlineClaimsAllowed=false`
  - `package-action-proof/provenance-boundary` remains missing and owned by release
  - package compact reports do not embed `inputSha256` or `outputBytes`
- Updated `research/experiments/syntheses/2026-05-release-claim-board.md` with a provenance-boundary matrix.

## Results

Local audit result:

| Field | Value |
| --- | --- |
| Release index signed | `false` |
| Release index attestation | `false` |
| Release gate | `blocked-by-publication-policy` |
| Headline claims allowed | `false` |
| Package provenance gate | `provenance-boundary(missing,release)` |
| Package boundary | local package-part evidence only; not signed provenance |
| Compact package boundary | not signed provenance, SLSA, in-toto, third-party attestation, or full streaming parity |
| Compact report embeds proof digests/artifact bytes | `false` |

Boundary matrix:

| Boundary | Allowed claim | Forbidden claim |
| --- | --- | --- |
| Local shape digest | proof artifacts are identified by local digest and stable shape | digest evidence is tamper-evident provenance |
| Signed provenance | no signed provenance is produced | SLSA provenance, signed proof bundle, certified build origin |
| in-toto statement | future attestation could map artifacts to a subject after release policy exists | this proof is an in-toto attestation |
| GitHub/Sigstore attestation | attestation is out of scope | GitHub/Sigstore verified or transparency-log backed |
| Publication policy | compact report commands are reproducibility pointers | compact report digests are publishable before storage/privacy/canonicalization policy |

## Confidence

High that current code keeps provenance claims fail-closed. Medium that release copy is safe, because the final wording still needs a release owner to approve the local-proof boundary.

## Fold-in decision

Promote to topic synthesis only. Keep `provenance-boundary` missing in the release proof index. Do not add package-action code. A future release owner may either approve the local-proof wording or implement a real attestation pipeline as a separate release/security effort.

## Next question

Can the package-action `streaming-matrix-boundary` gate be made approval-ready without expanding the writer surface, or does it require a broader streaming proof matrix before any streaming wording appears in release copy?
