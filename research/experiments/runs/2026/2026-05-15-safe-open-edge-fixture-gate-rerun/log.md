# Safe Open Edge Fixture Gate Rerun

## Question

Can the safe-open `public-edge-fixtures` blocker be closed by replacing generated signed/unknown structural packages with checked-in public binary fixtures?

## Hypothesis

No. The checked-in public fixture corpus likely still lacks signed and unknown-part workbook packages, so the next owner action should be product approval of disclosed generated structural packages or deliberate external fixture acquisition.

## External sources checked

- ECMA-376 identifies Open Packaging Conventions as the standard package layer for OOXML files: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- Microsoft OPC documentation describes package digital signatures and package relationships: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft Protected View documents trust/read-only opening for potentially unsafe files, which remains competitor contrast rather than Ascend's claim: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653

## Why this matters to Ascend

Safe unknown workbook opening is the top product-shaped handoff. The biggest fixture blocker is not another open-plan surface; it is whether generated structural edge packages are acceptable evidence for signed/unknown topology, or whether product waits for public binary fixtures.

## Probe/implementation

Ran the checked-in public fixture scan:

```bash
bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json
```

No production code changed.

## Results

- Scanned root: `fixtures/xlsx`
- Public fixtures scanned: 351
- Rejected protected fixtures: 2
  - `fixtures/xlsx/calamine/pass_protected.xlsx`
  - `fixtures/xlsx/poi/protected_passtika.xlsx`
- `signatureOrUnknownMatches=[]`
- `replacementStatus=no-public-binary-replacement-found`
- Boundary: this proves only that the checked-in public corpus has no replacement; it does not prove no suitable public workbook exists elsewhere.

## Confidence

High for the checked-in corpus result. Medium for the broader owner decision because external public signed/unknown workbooks may exist but still need licensing, stability, and vendoring review.

## Fold-in decision

Promote to topic synthesis and owner handoff. The next safe-open owner should either accept disclosed generated structural fixtures for guarded local proof or explicitly acquire and vendor public signed/unknown binary fixtures. Do not add another open-plan API.

## Next question

Can the package-action `edge-fixture-policy` blocker be resolved the same way: accept disclosed generated edge packages for structural package-proof wording, while forbidding trust/provenance claims?
