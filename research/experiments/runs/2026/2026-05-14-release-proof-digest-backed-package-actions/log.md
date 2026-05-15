# Release proof digest-backed package actions

## Question

Can the release proof bundle carry enough byte-level package evidence to support the product claim "auditable package-part mutation" without forcing normal agent plan/commit results to retain workbook bytes?

## Hypothesis

Yes. The existing package action proof can already attach optional per-part SHA-256 digests when callers provide source and output package bytes. The missing fold-in is to let `createReleaseProofBundle` pass those optional bytes into its plan and commit package action proofs.

## External sources checked

- [SLSA provenance v1.0-rc1](https://slsa.dev/spec/v1.0-rc1/provenance) defines artifacts as subjects with digest sets and materials, which supports keeping Ascend's claim digest-shaped rather than narrative-only.
- [Sigstore cosign blob signing](https://docs.sigstore.dev/cosign/signing/signing_with_blobs/) signs arbitrary blob artifacts and stores verification metadata in a bundle, reinforcing that Ascend's local proof bundle should be explicit about not being a signed attestation.
- [Sigstore quickstart with cosign](https://docs.sigstore.dev/quickstart/quickstart-cosign/) shows the common verify-blob flow, which is the right competitor/provenance contrast for any future signed proof bundle.
- [Microsoft Open Packaging Conventions overview](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview) describes XLSX as an OPC package of parts and relationships, which is the correct unit for part-level mutation evidence.
- [System.IO.Packaging docs](https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging) expose part-oriented package APIs, supporting package-part evidence as the honest abstraction boundary.

## Why this matters to Ascend

"Auditable package-part mutation" is one of the top claim-ladder candidates. A release proof that only says plan/commit/reopen passed is useful, but it does not show which package parts have byte evidence. Digest-backed package action proofs make the claim inspectable: preserved/generated parts can carry source/output hashes, equality can be checked, and missing coverage is visible in the coverage counters.

## Probe/implementation

- Inspected `packages/sdk/src/agent-workflow.ts` and confirmed `createPackageActionProof` already supports optional `sourceBytes` and `outputBytes`.
- Confirmed `createReleaseProofBundle` created plan/commit package action proofs without passing bytes.
- Added optional `sourceBytes` and `outputBytes` to `ReleaseProofBundleOptions`.
- Passed `sourceBytes` into the plan package action proof.
- Passed both `sourceBytes` and `outputBytes` into the commit package action proof.
- Extended the release proof bundle test to pass real fixture bytes and assert source/output digest coverage plus per-action output hashes.

## Results

- `bun test packages/sdk/src/agent-workflow.test.ts -t "release proof bundle"` passed.
- `bun test packages/sdk/src/agent-workflow.test.ts -t "package action proof"` passed.
- Normal plan and commit surfaces remain unchanged; byte retention is opt-in at proof-bundle construction time.
- The bundle now exposes digest coverage counters:
  - plan source digest count
  - commit source digest count
  - commit output digest count
  - matching and mismatched per-part digest counts

## Confidence

High for the scoped claim that release proof bundles can include opt-in digest-backed package action evidence. Medium for the broader product claim because it still needs a real-file proof fixture and CLI/MCP surfacing before it is easy for users to run.

## Fold-in decision

Promote to correctness loop and product/DX loop. The SDK helper now has the needed proof hook; the next implementation loop should expose this as a release-proof CLI or MCP artifact with explicit claim boundaries.

## Next question

What is the smallest real-workbook release proof fixture and CLI/MCP command that demonstrates inspect, plan, commit, reopen, diff, package-action digest coverage, and audit boundaries without embedding private workbook content?
