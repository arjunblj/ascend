# Package Action Proof Source Graph

## Question

Can commit-local package action proofs distinguish newly added package parts from regenerated existing parts without adding a product surface or changing writer behavior?

## Hypothesis

Yes. Commit workflows already retain source and output package bytes for optional digest-backed package action proofs. If the proof helper infers the source package graph from those source bytes, it can classify generated output parts as `add` when the part was absent from the input package and `regenerate` when it existed.

## External sources checked

- Microsoft `System.IO.Packaging` documents packages as parts plus package-level and part-level relationships: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging
- Microsoft `Package` docs include package part and relationship create/delete APIs, supporting part-level mutation accounting: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging.package
- SheetJS CE write options document writer scope, data-preservation orientation, and explicit VBA preservation knobs: https://docs.sheetjs.com/docs/api/write-options/
- openpyxl tutorial documents `keep_vba` and warns unsupported workbook objects may be lost on save: https://openpyxl.readthedocs.io/en/stable/tutorial.html

## Why this matters to Ascend

The release claim board ranks "auditable package-part mutation" as the next top correctness/product claim after safe open. Ascend already exposes a stable action taxonomy, boundaries, digest evidence, and opt-in CLI/API/MCP package action proofs. But commit proofs using only byte archives still lacked source graph membership, so a new worksheet part created by `addSheet` was classified as `regenerate` with the reason "source graph was not provided." That weakens the product-shaped claim because `add` versus `regenerate` is part of the promised proof vocabulary.

## Probe/implementation

- Inspected `createPackageActionProof()`, `createAgentCommitPackageActionProof()`, CLI/API/MCP package action surfaces, release claim board, and package action proof tests.
- Ran a local add-sheet commit probe before the change:

```json
{
  "byAction": { "passthrough": 3, "regenerate": 6, "add": 0, "drop": 0, "error": 0 },
  "sourceGraphIncluded": false,
  "addExamples": []
}
```

- Folded in a small SDK helper change:
  - when `createPackageActionProof()` receives `sourceBytes` but no explicit `sourcePackageGraph`, it infers the source graph with `inspectXlsxPackageGraph(sourceBytes)`;
  - coverage now marks `sourceGraphIncluded: true` and records source part/relationship counts for commit-local proofs;
  - generated parts absent from the inferred source graph classify as `add`.
- Updated the focused package action proof test to commit `addSheet` and assert add classification, source graph coverage, source relationship counts, byte digest coverage, and `sourcePresent: false` on the added worksheet part.

## Results

After the fold-in, the same add-sheet commit probe reports:

```json
{
  "byAction": { "passthrough": 3, "regenerate": 5, "add": 1, "drop": 0, "error": 0 },
  "sourceGraphIncluded": true,
  "sourcePartCount": 8,
  "sourceRelationshipCount": 5,
  "addExamples": [
    {
      "action": "add",
      "partPath": "xl/worksheets/sheet2.xml",
      "sourcePresent": false
    }
  ]
}
```

Validation passed:

- `bun test packages/sdk/src/agent-workflow.test.ts -t "package action proof"`
- `bunx biome check packages/sdk/src/agent-workflow.ts packages/sdk/src/agent-workflow.test.ts`
- `bun test apps/cli/src/cli.test.ts -t "safe agent workflow"`
- `bun test apps/api/api.test.ts -t "safe write workflow"`
- `bun test apps/mcp/src/index.test.ts -t "package action proof"`
- `bunx tsc --build`
- `bunx biome check`
- `bun run test:changed` (4997 pass, 1 skip)

## Confidence

High for the scoped fold-in. It is additive evidence derivation over bytes the commit workflow already retains, and it does not change writer output or default response payloads. Medium for the broader release claim because the proof still needs a public workflow bundle covering docProps passthrough, calc-chain drop, signature invalidation, macro/ActiveX preservation, drawing/chart sidecars, and unknown part rejection.

## Fold-in decision

Fold into production as a small correctness/product proof improvement. This closes the commit-local `add` versus `regenerate` evidence gap, but it does not promote a new API/CLI/MCP surface and does not make the full auditable package-part mutation claim release-ready by itself.

## Next question

Can a tracked package action proof harness generate a fixture-backed report for passthrough, regenerate, add, drop, and error cases without expanding default product payloads?
