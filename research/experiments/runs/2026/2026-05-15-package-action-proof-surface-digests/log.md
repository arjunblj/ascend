# Package Action Proof Surface Digests

## Question

Can the opt-in CLI/API/MCP package action proof surfaces include per-part source/output digest evidence for commits without bloating default commit responses or claiming signed provenance?

## Hypothesis

Ascend already computes source and output bytes during commit verification. If those bytes can be retained only in process and consumed only by the opt-in package action proof helper, the "auditable package-part mutation" claim becomes materially stronger with a small implementation.

## External sources checked

- Microsoft Open Packaging Conventions fundamentals: packages are graphs of parts and relationships, so package-part mutation claims should be scoped to package parts and relationship evidence. https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft `System.IO.Packaging` docs: OPC packages have package parts, relationships, and optional digital signatures; Ascend's local evidence is not equivalent to `PackageDigitalSignature`. https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging?view=windowsdesktop-10.0
- SLSA provenance spec: stronger provenance models identify subjects/materials with digest sets and require downstream verification of untrusted provenance. https://slsa.dev/spec/v1.0-rc1/provenance
- in-toto attestation model via SLSA/GitHub artifact attestation docs: signed artifact attestations are a separate envelope from local digest evidence. https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

"Auditable package-part mutation" is one of the top claim-ladder candidates. Before this fold-in, CLI/API/MCP packageActions named passthrough/regenerate/add/drop/error actions but exposed zero source/output byte digest coverage on commit. That made the proof useful for explanation but weak for audit: a caller could not see which package parts had source/output SHA-256 evidence or byte-equality classification.

## Probe/implementation

Local probe: `probes/surface-digests.ts` copies `fixtures/xlsx/poi/SampleSS.xlsx`, runs `ascend plan --package-actions --json`, then runs `ascend commit --package-actions --json`, and summarizes proof coverage.

Initial probe result:

- plan `sourceByteDigestCount`: 0
- plan `outputByteDigestCount`: 0
- commit `sourceByteDigestCount`: 0
- commit `outputByteDigestCount`: 0
- commit actions with `outputSha256`: 0
- commit actions with `bytesEqual`: 0

Fold-in:

- Added `createAgentCommitPackageActionProof(result)` in `packages/sdk/src/agent-workflow.ts`.
- Stored commit-local source/output bytes in an SDK-private `WeakMap`, keyed by the live `AgentCommitResult`, so JSON serialization of ordinary commit results does not include workbook bytes.
- Switched opt-in commit package action surfaces in CLI/API/MCP to call the commit helper.
- Added focused SDK, CLI, API, and MCP tests that assert opt-in commit proofs include source/output digest coverage and at least one action with `outputSha256`.

Post-change probe result:

- plan `sourceByteDigestCount`: 0
- plan `outputByteDigestCount`: 0
- commit `sourceByteDigestCount`: 12
- commit `outputByteDigestCount`: 12
- commit `matchingByteDigestCount`: 9
- commit `mismatchedByteDigestCount`: 3
- commit actions with `outputSha256`: 12
- commit actions with `bytesEqual`: 12

## Results

The fold-in strengthens the production claim from "Ascend can explain package actions" to "Ascend can provide opt-in local package-part digest evidence for committed package actions." This is still not signed provenance and not Excel semantic equivalence, but it is enough to support an honest audit surface for package-part mutation.

Validation:

- `bun test packages/sdk/src/agent-workflow.test.ts -t "commit package action proof uses commit-local byte evidence"`
- `bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow"`
- `bun test apps/api/api.test.ts -t "plan and commit endpoints provide the safe write workflow"`
- `bun test apps/mcp/src/index.test.ts -t "ascend.plan and ascend.commit can include package action proof evidence"`
- `bunx biome check packages/sdk/src/agent-workflow.ts packages/sdk/src/index.ts packages/sdk/src/agent-workflow.test.ts apps/cli/src/commands/commit.ts apps/cli/src/cli.test.ts apps/api/src/server.ts apps/api/api.test.ts apps/mcp/src/index.ts apps/mcp/src/index.test.ts`
- `bunx tsc --build`
- `bun run test:changed`

## Confidence

High for non-destructive commit paths and prepared/direct commit surfaces using live SDK results. Medium for any downstream caller that serializes/deserializes commit results before asking for package action proof; that path intentionally cannot recover byte evidence from the private WeakMap.

## Fold-in decision

Promote to correctness loop and product/DX loop. This is now implemented as a small production fold-in with focused validation.

Honest boundary: the proof remains local evidence. It does not sign the workbook, attest the runtime, prove third-party provenance, or prove Excel recalculation equivalence. Plan packageActions still omit output digests because there is no committed output package yet.

## Next question

Can the release proof bundle CLI/API/MCP handoff expose a compact proof artifact path with the same digest-backed package actions, without embedding private workbook bytes or creating fake attestation claims?
