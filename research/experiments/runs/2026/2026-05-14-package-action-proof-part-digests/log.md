# Package Action Proof Part Digests

## Question

Can package action proofs attach optional per-part byte digests without changing writer behavior or claiming signed provenance?

## Hypothesis

Yes. The existing proof helper can accept source/output package bytes, hash the package part behind each action, and report whether both byte hashes match. This makes passthrough evidence stronger for SDK callers while keeping CLI/API/MCP unchanged until a release-proof owner wires bytes into those surfaces deliberately.

## External sources checked

- In-toto/Sigstore attestation shape via Cosign blob signing: https://docs.sigstore.dev/cosign/signing/signing_with_blobs/
- Sigstore quickstart and attestation framing: https://docs.sigstore.dev/quickstart/quickstart-cosign/
- SLSA provenance digest guidance: https://slsa.dev/spec/v1.0-rc1/provenance
- Microsoft Open Packaging Conventions fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft `System.IO.Packaging` package relationships and signatures: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging

## Why this matters to Ascend

The claim ladder says auditable package-part mutation needs part-level proof, not just action labels. Digest evidence is the smallest useful next step: it lets a proof say which source/output package part bytes were observed and whether they match, while avoiding unsupported claims about signatures, external attestations, or relationship-level action closure.

## Probe/implementation

- Inspected `createPackageActionProof()` and the package graph byte-preservation audit.
- Added optional `sourceBytes` and `outputBytes` to `PackageActionProofOptions`.
- Added optional `sourceSha256`, `outputSha256`, and `bytesEqual` to `PackageActionProofEntry`.
- Extended proof coverage with source/output digest counts and matching/mismatched digest counts.
- Kept digest evidence opt-in and local to the SDK helper. Existing CLI/API/MCP package-action surfaces are unchanged unless a future owner deliberately passes package bytes.
- Added a fixture-backed SDK test that proves matching and mismatched preserved package part digests.

## Results

Focused validation passed:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "package action proof"
bun test packages/sdk/src/agent-workflow.test.ts -t "release proof bundle"
bunx biome check packages/sdk/src/agent-workflow.ts packages/sdk/src/agent-workflow.test.ts packages/sdk/src/index.ts
bunx tsc --build
bun run test:changed
```

The new test builds small OPC/XLSX packages with a preserved `custom/item.xml` part and verifies:

- matching source/output part bytes yield `bytesEqual: true`;
- changed output bytes yield `bytesEqual: false`;
- proof coverage reports source/output digest counts and matching/mismatched counts.

`bun run test:changed` initially found four unrelated-looking CLI failures that passed when rerun directly. The full gate then passed on rerun with 4979 tests, 1 skip, and 0 failures.

## Confidence

Medium-high for the helper behavior. It is additive, opt-in, and does not affect writer behavior. It is not yet a complete product proof because the normal plan/commit surfaces do not pass package bytes, and relationship-level actions are still summarized through audits rather than per-relationship action entries.

## Fold-in decision

Folded into the correctness loop as an SDK proof primitive. Do not promote this as signed provenance. The next owner should be the release-proof/product loop, which can decide whether to pass source/output bytes into full proof bundles while keeping compact surfaces bounded.

## Next question

Can a release proof bundle include digest-backed package action evidence without making CLI/API/MCP compact responses too large?
