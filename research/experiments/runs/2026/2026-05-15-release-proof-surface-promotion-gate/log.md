# Release Proof Surface Promotion Gate

## Question

Should Ascend promote release proof bundles from SDK-only evidence into CLI/API/MCP surfaces now, or hold until the product loop defines artifact storage, privacy, and verification semantics?

## Hypothesis

Hold for product/DX design. The SDK proof bundle is strong enough to support the product-shaped claim, but a CLI/API/MCP surface that returns or writes proof artifacts needs explicit policy for artifact paths, private workbook data, compactness, and honest non-attestation wording.

## External sources checked

- GitHub artifact attestations verify artifact integrity and provenance through cryptographically signed attestations, and `gh attestation verify --format=json` emits verified attestation JSON. https://cli.github.com/manual/gh_attestation_verify
- GitHub Actions artifact attestation docs distinguish generating attestations from verifying them later with GitHub CLI. https://docs.github.com/actions/security-guides/using-artifact-attestations-to-establish-provenance-for-builds
- Sigstore bundle docs define bundles as verification material plus signature content; this is stronger than Ascend's current local proof JSON. https://docs.sigstore.dev/about/bundle/
- Sigstore overview describes verification against identity, roots of trust, and transparency log inclusion. https://docs.sigstore.dev/
- SLSA provenance frames provenance as subject artifacts, materials, digests, and downstream verification. https://slsa.dev/spec/v1.0-rc1/provenance

## Why this matters to Ascend

"Release proof bundle" is product-shaped: prove inspect/plan/commit/reopen/diff/audit on a real workbook without fake claims. Exposing it prematurely could make Ascend look more trustworthy while actually weakening the trust boundary, especially if users mistake local evidence for signed provenance or if private workbook material leaks into saved proof artifacts.

## Probe/implementation

Inspected current implementation:

- `packages/sdk/src/agent-workflow.ts` exports `createReleaseProofBundle(plan, commit, options)`.
- The bundle includes subject digests, plan/commit trace digests, operation artifact digests, reopen status, audit booleans, diff summary, package action proofs, consistency checks, and claim boundaries.
- CLI/API/MCP expose plan, commit, packageActions, and release-adjacent audit evidence, but no first-class release proof artifact command/endpoint/tool.
- `apps/cli/src/index.ts` has no proof-bundle command and commit has no `--proof-bundle` flag.
- `apps/api/src/server.ts` and `apps/mcp/src/index.ts` have no release-proof route/tool.

Validation/probe:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "release proof bundle links plan, commit, reopen, diff, and audit evidence"
```

The existing fixture-backed SDK test passed with 21 assertions. It proves the SDK helper works, including package action digest evidence and honest claim boundaries.

## Results

Do not fold in a CLI/API/MCP proof surface in this loop.

Reasons:

- The SDK proof exists, but surface semantics are unresolved: return inline JSON, write a sidecar file, store in a handle, or produce a compact summary plus artifact path.
- Direct commit flows do not always retain an explicit `AgentPlanResult`; recomputing a plan inside a proof surface could duplicate work and confuse the hash/plan-digest story.
- A signed-provenance shaped interface would be misleading until Ascend has an external attestation envelope, verifier roots, or at least a documented "local evidence only" artifact format.
- Private workbook safety is a product decision. The proof should never embed workbook bytes by default, and proof artifact paths should be explicit.

Recommended product-loop handoff:

1. Add a CLI-only proof artifact first, not all surfaces at once: `ascend release-proof <file> --ops ops.json --output out.xlsx --proof proof.json --expect-sha256 <hash>`.
2. Make the command run plan and commit in one process so it can pass source/output bytes to `createReleaseProofBundle`.
3. Require explicit `--proof <path>`; do not print full proof JSON by default.
4. Default stdout to a compact summary with `proofPath`, `inputSha256`, `outputSha256`, `planDigest`, `consistency.valid`, and claim boundaries.
5. Add `--attestation none` as the only initial mode if the product loop wants a future extension point.
6. After CLI behavior stabilizes, mirror it in API/MCP as a prepared-plan based workflow.

## Confidence

High that the SDK proof bundle is production-useful. Medium that the exact CLI shape above is final. The right next step is a product/DX owner decision, not another opportunistic surface patch.

## Fold-in decision

Promote to product/DX loop only. Do not promote to correctness or performance implementation yet.

## Next question

Can formula language-service primitives make a credible "formula intelligence" claim today, or do rename/code actions still need one more correctness fold-in before promotion?
