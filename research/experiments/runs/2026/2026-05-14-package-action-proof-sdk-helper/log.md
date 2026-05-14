# Package Action Proof SDK Helper

## Question

Can Ascend turn the existing XLSX write plan into a proof artifact that explains each package part as `passthrough`, `regenerate`, `add`, `drop`, or `error` without changing writer behavior?

## Hypothesis

Yes. `writePlanSummary`, `writePolicy`, and package graph audits already contain enough evidence to classify package actions for agent review. A small SDK helper can make that evidence explicit and testable before investing in deeper writer changes.

## External sources checked

- Microsoft Open Packaging Conventions fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- ECMA-376 Office Open XML file formats, Part 2 Open Packaging Conventions: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- Microsoft Open XML SDK package part guidance: https://learn.microsoft.com/en-us/office/open-xml/general/how-to-add-a-new-document-part-to-a-package
- Microsoft `PackageDigitalSignature` documentation: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging.packagedigitalsignature

## Why this matters to Ascend

Preservation-first XLSX mutation needs a user- and agent-readable explanation of what happened to package content. OPC treats a workbook as a graph of parts and relationships, and signatures cover package parts/relationships, so an explainable package action proof is a better audit unit than a raw list of generated XML files.

## Probe/implementation

- Inspected Ascend's current `summarizePlannedWrite`, `AscendWorkbook.writePlanSummary`, `WritePolicyReport`, and package graph audit helpers.
- Added `createPackageActionProof()` to `packages/sdk/src/agent-workflow.ts`.
- Exported the helper and proof types from `packages/sdk/src/index.ts`.
- Added targeted tests in `packages/sdk/src/agent-workflow.test.ts`:
  - a real workbook add-sheet plan separates source-backed regenerations from new package additions;
  - a synthetic proof records passthrough, drop, and error evidence.

## Results

The helper produces `ascend-package-action-proof` with stable action counts, per-part entries, optional source package graph evidence, write-policy diagnostic codes, package graph audit issue codes, and error issue strings.

Targeted validation passed:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "package action proof"
```

The proof currently remains local evidence. It explains planned package actions but does not yet prove relationship-level closure, byte hashes for every part, or signed external attestations.

## Confidence

High for making existing write-plan evidence easier to audit. Medium for full package proof semantics because relationship action classification and per-part digest evidence should be added in a later fold-in.

## Fold-in decision

Promote to correctness loop and product/DX loop. This cycle folded in the first production helper and tests. Next fold-in should attach this proof to release proof bundles and add relationship-level action coverage.

## Next question

Can the existing compact agent view be made budget-aware enough to produce stable, token-bounded workbook summaries for CLI/MCP agents?
