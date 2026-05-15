# Package Action Proof Coverage Boundaries

## Question

Can the existing package action proof make its coverage boundaries explicit enough to support the product claim "auditable package-part mutation" without adding a new surface or implying full relationship/hash proof?

## Hypothesis

Yes. The package action proof already classifies package parts as `passthrough`, `regenerate`, `add`, `drop`, or `error`. A tiny fold-in can add coverage metadata that says which evidence inputs were present and whether relationship or byte-preservation audit issues exist, while still making clear that the action taxonomy is package-part scoped.

## External sources checked

- Microsoft Open Packaging Conventions fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- ECMA-376 Office Open XML, Part 2 Open Packaging Conventions: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- Microsoft `System.IO.Packaging` relationship and signature model: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging
- openpyxl usage warning about unsupported item loss on save: https://openpyxl.readthedocs.io/en/3.0/usage.html
- SheetJS write options and VBA preservation controls: https://docs.sheetjs.com/docs/api/write-options/
- SheetJS VBA/macro blob behavior: https://docs.sheetjs.com/docs/csf/features/vba/

## Why this matters to Ascend

The claim ladder ranks auditable package-part mutation as the top correctness/product proof. OPC packages are parts plus relationships, and signatures can cover parts and relationships. If Ascend exposes only part actions, agents may overread the proof as relationship closure or per-part digest evidence. Coverage metadata keeps the claim honest while still making the existing proof more useful.

## Probe/implementation

- Inspected `createPackageActionProof()` and `createReleaseProofBundle()` in `packages/sdk/src/agent-workflow.ts`.
- Inspected existing package graph fidelity audits for preserved relationship and byte-preservation issue codes.
- Added `PackageActionProofCoverage` to the SDK proof object:
  - declares `proofScope: "package-part-actions-with-audit-summaries"`;
  - records whether source package graph, write policy, and package graph audit evidence were included;
  - records source part and relationship counts when a source graph is provided;
  - records relationship-audit and byte-preservation-audit issue counts.
- Exported the new coverage type.
- Added targeted tests proving coverage on a real package action proof and synthetic relationship/byte audit issues.

## Results

Focused validation passed:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "package action proof"
bun test packages/sdk/src/agent-workflow.test.ts -t "release proof bundle"
bunx biome check packages/sdk/src/agent-workflow.ts packages/sdk/src/agent-workflow.test.ts packages/sdk/src/index.ts research/experiments/index.md research/experiments/runs/2026/2026-05-14-package-action-proof-coverage-boundaries/log.md
bunx tsc --build
bun run test:changed
```

The proof now exposes enough metadata for agents to distinguish:

- package-part action classification;
- source graph evidence present or absent;
- write-policy evidence present or absent;
- package-graph audit evidence present or absent;
- relationship audit issue counts;
- byte-preservation audit issue counts.

`bun run test:changed` expanded to the full suite and passed with 4974 tests, 1 skip, and 0 failures.

## Confidence

High for the narrow fold-in. It is additive metadata on the existing opt-in proof object and does not change writer behavior. Medium for the broader product claim, because per-relationship actions and per-part source/output digests are still future proof requirements.

## Fold-in decision

Folded into the correctness and product/DX loops as a tiny proof-boundary improvement for the top claim. The owner of the next deeper implementation should still be the correctness loop, not opportunistic research.

## Next question

Can the safe unknown workbook opening proof bundle measure open-plan latency versus full hydration across public fixtures without adding another product surface?
