# Release Proof Ready-When Checklists

## Question

Can the release proof index expose machine-readable `readyWhen` checklists so the top claim owner loops know exactly which blocker would flip `headlineClaimAllowed`, without promoting another production surface?

## Hypothesis

Yes. The current index already carries commands, publication blockers, fixture provenance, digests, and headline gates. A small readiness checklist on each top artifact can turn broad claim stewardship into owner-specific exit criteria while staying in benchmark/proof code.

## External sources checked

- SLSA source verification expectations: https://slsa.dev/spec/v1.2/verifying-source
- GitHub artifact attestation provenance and `subject-digest` verification: https://docs.github.com/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
- in-toto attestation subjects and digests: https://github.com/in-toto/attestation

## Why this matters to Ascend

The ranked portfolio says the top two claims are safe unknown workbook opening and auditable package-part mutation. Both have local proof, but neither should become headline release copy until publication blockers are resolved or explicitly accepted. Machine-readable readiness requirements prevent research from sliding into new product surfaces and give product, performance, correctness, and release owners concrete proof gates.

## Probe/implementation

- Inspected `fixtures/benchmarks/release-proof-index.ts` and its test.
- Added `ReleaseProofReadinessRequirement` with `id`, `status`, `ownerLoop`, `requirement`, and optional `evidence`.
- Added `readyWhen` checklists for:
  - `safe-open-proof`: public/generated edge fixture policy, release latency run, and publication boundary approval.
  - `package-action-proof`: edge fixture policy, provenance boundary, and unsupported-feature boundary approval.
- Rendered the checklist in the Markdown index.
- Kept both artifacts at `headlineClaimAllowed: false` and `releaseGate: blocked-by-publication-policy`.

## Results

Current no-timings proof index:

| Artifact | Ready-when IDs | Stable shape SHA-256 | Summary |
| --- | --- | --- | --- |
| safe-open-proof | public-edge-fixtures, release-latency-run, publication-boundary | `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178` | cases=9, ok=8, rejected=1, reviewBeforeHydration=4, malformedRejected=true |
| package-action-proof | edge-fixture-policy, provenance-boundary, unsupported-feature-boundary | `60da8baa4a897e7edbd3f02fcb1a7026643bc68de1f5dedec082e71d29f03213` | cases=8, passthrough=27, regenerate=38, add=3, drop=3, error=1 |

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

Focused test result: 3 pass, 0 fail.

## Confidence

High for the narrow fold-in. The change is additive proof metadata, tested, and does not alter writer/opening behavior or create a user-facing product surface. Medium for publication policy because product/release owners still need to decide whether disclosed generated edge packages are acceptable for headline proof.

## Fold-in decision

Fold into the release proof harness only. This should be handed to the product/performance and correctness/product owner loops as the top two claim gates. Do not promote formula rename, columnar sidecars, or new proof surfaces from this change.

## Next question

Can the top two owner prompts consume `readyWhen` directly and produce release reports that fail closed when any requirement remains `missing`?
