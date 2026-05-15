# Safe Open External Candidate License Refresh

## Question

Does the previously found ExcelForge workbook improve the safe-open public-edge-fixture gate now that license metadata is clearer?

## Hypothesis

It improves the unknown-part fixture acquisition story but does not close the `public-edge-fixtures` gate. The candidate still is not vendored into Ascend, and it does not solve the signed-workbook case.

## External sources checked

- ExcelForge repository page, which currently shows an MIT license badge and documents digital-signature support: https://github.com/node-projects/excelForge
- ExcelForge package manifest, which declares `"license": "MIT"`: https://raw.githubusercontent.com/node-projects/excelForge/master/package.json
- ExcelForge raw candidate workbook: https://raw.githubusercontent.com/node-projects/excelForge/master/src/test/Book%201.xlsx
- Microsoft Open Packaging Conventions overview for package parts and relationships: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

Safe unknown workbook opening is a top release-claim candidate, but its product gate remains blocked on generated signed/unknown structural fixtures. A license-clear public unknown-part workbook would reduce one part of that blocker without adding a new open surface.

## Probe/implementation

- Downloaded the raw candidate workbook to `/private/tmp/excelForge-Book1.xlsx`.
- Downloaded the raw package manifest to `/private/tmp/excelForge-package.json`.
- Computed SHA-256 digests:
  - workbook: `9c5426fa71ff68cc7e40e19e02b5992daf91da5754ef643d2db2f89bd70bb122`;
  - package manifest: `cae1feec581eed864255cff45fa23a7e2c085cb0f2c2628d1a0187fc39de3ef7`.
- Ran `inspectWorkbookOpenPlan` against the workbook.
- Parsed the package manifest license field.

## Results

- Manifest evidence now reports:
  - package name `@node-projects/excelforge`;
  - version `3.6.0`;
  - license `MIT`;
  - repository `git+https://github.com/node-projects/excelForge.git`.
- Safe-open plan for `Book 1.xlsx`:
  - `recommendedMode=metadata-only`;
  - `reviewBeforeHydration=true`;
  - `partCount=50`;
  - `relationshipCount=37`;
  - risk family `preservedOther`;
  - unknown sample part `docMetadata/LabelInfo.xml`.
- This upgrades the candidate from "license metadata unclear" to "candidate has MIT package-manifest evidence." It still does not close the release gate because the workbook is not checked in here, release owners have not approved vendoring policy, and the signed-workbook fixture gap remains.

## Confidence

Medium. The local probe is repeatable and the package manifest is explicit, but vendoring a binary fixture still needs product/release policy approval and ideally a repository-level license file or owner-approved attribution rule.

## Fold-in decision

Promote to topic synthesis and owner handoff. Do not vendor the workbook or mark `public-edge-fixtures` satisfied in research.

## Next question

Should product/release accept the ExcelForge workbook as the public unknown-part replacement while separately keeping the signed-workbook fixture gate open?
