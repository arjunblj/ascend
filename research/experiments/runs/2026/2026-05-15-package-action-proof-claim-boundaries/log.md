# Package action proof claim boundaries

## Question

Can auditable package-part mutation take a small production step toward a product-shaped proof claim without changing writer behavior or implying signed provenance?

## Hypothesis

Yes. The package action proof already exposes the core action taxonomy: passthrough, regenerate, add, drop, and error. The missing product-safety gap is that the proof object itself does not carry honest claim boundaries, so CLI/API/MCP consumers can receive `ascend-package-action-proof` JSON without seeing what the proof does not mean.

## External sources checked

- Microsoft Open Packaging Conventions fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft `System.IO.Packaging` namespace: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging
- openpyxl tutorial and preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options: https://docs.sheetjs.com/docs/api/write-options/
- SLSA Verification Summary Attestation: https://slsa.dev/verification_summary/v1

## Why this matters to Ascend

The product claim is "auditable package-part mutation," not merely "writer emits XML." OPC packages are graph-shaped, signatures are separate validation artifacts, and mature provenance systems separate local evidence from signed attestations. Ascend should carry the same boundary directly in its proof JSON so agents can distinguish local plan/audit evidence from cryptographic provenance or Excel semantic certification.

## Probe/implementation

- Inspected `createPackageActionProof()` in `packages/sdk/src/agent-workflow.ts`.
- Confirmed the existing proof schema already has:
  - `PackageActionKind = 'passthrough' | 'regenerate' | 'add' | 'drop' | 'error'`;
  - per-action reasons;
  - coverage counters;
  - optional source/output part digests and `bytesEqual`;
  - issue extraction from blocker diagnostics and package graph audit issues.
- Confirmed current tests already exercise all five actions at SDK level.
- Added `claimBoundaries` to `PackageActionProof` and `claimBoundaries` override support to `PackageActionProofOptions`.
- Added default package-action proof boundaries:
  - local package-part evidence, not signed provenance or third-party attestation;
  - write-plan and audit evidence, not Excel semantic recalculation equivalence;
  - passthrough byte equality only when source/output digests are present and `bytesEqual` is true;
  - drop/error actions require caller review before claiming preservation.
- Updated package action proof tests to assert boundary presence.

## Results

Production fold-in:

- `packages/sdk/src/agent-workflow.ts`
- `packages/sdk/src/agent-workflow.test.ts`

Focused validation passed:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "package action proof"
bunx biome check packages/sdk/src/agent-workflow.ts packages/sdk/src/agent-workflow.test.ts
bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow"
bun test apps/api/api.test.ts -t "plan and commit endpoints provide the safe write workflow"
bun test apps/mcp/src/index.test.ts -t "package action proof"
bunx tsc --build
```

`bun run test:changed` ran the full suite and failed in dirty, pre-existing journal work outside this cycle: `packages/sdk/src/journal-exactness.test.ts` expected a legacy array edit to apply, but current dirty journal changes block the edit with `Cannot edit A1 because it is part of legacy array formula A1:A2`. This cycle did not touch `packages/sdk/src/journal.ts`, `packages/sdk/src/journal-exactness.test.ts`, or `packages/sdk/src/interactive-contract.test.ts`.

The proof schema now carries its own honest boundary text everywhere `createPackageActionProof()` is returned, including existing SDK/CLI/API/MCP package action proof surfaces.

## Confidence

High for the boundary fold-in. It is additive, does not change writer behavior, and is covered by focused tests. Medium for the broader product claim, because a public fixture-backed workflow bundle still needs to prove real docProps passthrough, generated worksheet XML, calc-chain drop, signature invalidation, macro/ActiveX preservation, drawing/chart sidecars, and unknown part rejection.

## Fold-in decision

Fold into production as a tiny correctness/product proof-safety improvement. Do not promote the full "auditable package-part mutation" claim yet; promote only "package action proofs include explicit local-evidence boundaries."

## Next question

Can a public fixture workflow bundle prove package actions across passthrough, regenerate, add, drop, and error with digest evidence, without bloating compact CLI/API/MCP responses?
