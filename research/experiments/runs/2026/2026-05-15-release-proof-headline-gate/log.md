# Release Proof Headline Gate

## Question

Should the generated-edge fixture policy be enforced by the release proof index itself?

## Hypothesis

Yes. The release proof index is the machine-readable artifact future release work will consult. If it only lists blockers as text, a release loop could still overread local proof as headline-ready. Adding an explicit headline gate makes the boundary enforceable.

## External sources checked

- SLSA software attestations: https://slsa.dev/spec/v1.1/attestation-model
- in-toto Attestation Framework: https://github.com/in-toto/attestation
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653

## Why this matters to Ascend

Ascend can have strong local proof without having a public headline claim. SLSA and in-toto make clear that real attestations are authenticated metadata with verification semantics; Ascend's current release proof index is not that. The index should therefore say when evidence is local-proof-only and blocked from headline copy.

## Probe/implementation

Updated `fixtures/benchmarks/release-proof-index.ts`:

- Added `headlineClaimAllowed`.
- Added `releaseGate`.
- Marked both current top artifacts as `headlineClaimAllowed: false`.
- Marked both current top artifacts as `blocked-by-publication-policy`.

The release claim board now records that this gate is intentional: unresolved blockers require product acceptance or fixture replacement before stronger release copy.

## Results

Current top artifacts:

| Artifact | Release gate | Headline claim allowed | Why |
| --- | --- | --- | --- |
| `safe-open-proof` | `blocked-by-publication-policy` | false | signed/unknown edge cases are generated, and timings are local proof data |
| `package-action-proof` | `blocked-by-publication-policy` | false | synthetic edge packages must stay disclosed, and proof is not signed attestation |

The gate keeps the release proof index useful for local evidence while preventing accidental publication wording.

## Confidence

High. This is a small benchmark-harness metadata fold-in with focused tests, and it matches the release fixture policy.

## Fold-in decision

Promote to product/correctness proof packaging. Do not add production surfaces.

## Next question

Can a future release owner resolve the gates by accepting generated edge proof under the fixture policy, or by adding clear-license public binary fixtures?
