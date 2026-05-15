# Safe Open External Fixture Candidate Scan

## Question

Can the safe unknown workbook opening proof replace any generated edge package with a durable public workbook fixture instead of relying only on disclosed synthetic topology fixtures?

## Hypothesis

Maybe for the unknown-part case, but not for the signed-workbook case. A constrained public scan may find real XLSX files with unknown package parts, while signed XLSX files are still rare and hard to vendor safely.

## External sources checked

- Microsoft Protected View frames unsafe files as a trust/review UX, not proof that active content is safe: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft OPC fundamentals define packages as parts and relationships: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Stack Overflow discussion of signed XLSX package additions mentions `_xmlsignatures/origin.sigs` and `sig1.xml`, useful as search terms rather than a fixture source: https://stackoverflow.com/questions/21387656/how-to-add-a-digital-signature-to-an-xlsx-file
- node-projects/excelForge documents OOXML signing APIs and includes public test XLSX/XLSM files: https://github.com/node-projects/excelForge
- GitHub repository API for `node-projects/excelForge` reported `license: null`: https://api.github.com/repos/node-projects/excelForge
- Raw public candidate checked by probe: https://raw.githubusercontent.com/node-projects/excelForge/master/src/test/Book%201.xlsx

## Why this matters to Ascend

`safe-open-proof/public-edge-fixtures` is the first product blocker for Ascend's top release claim. If a real public workbook can replace a generated unknown-part package, the claim becomes less synthetic. If fixture provenance or licensing is weak, the release gate should stay blocked.

## Probe/implementation

Ran the checked-in public fixture scan:

```bash
bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json
```

Then ran a constrained external probe over public `node-projects/excelForge` test workbooks discovered through the GitHub API. The probe downloaded candidates to `/tmp`, ran `inspectWorkbookOpenPlan`, recorded risk families, and removed the temporary files.

Candidate files probed:

- `Book 1.xlsx`
- `Book 2.xlsx`
- `Book 3.xlsx`
- `ErrorsAndWarnings.xlsx`
- `allelements.xlsm`
- `userformvba.xlsm`

## Results

Checked-in fixture scan remains unchanged:

- scanned 351 checked-in XLSX/XLSM fixtures;
- rejected 2 protected fixtures;
- found 0 signed/unknown replacement candidates.

External candidate scan:

| Candidate | Result |
| --- | --- |
| `Book 1.xlsx` | `metadata-only`, `reviewBeforeHydration=true`, risk family `preservedOther`, part count 50, relationship count 37, unknown sample part `docMetadata/LabelInfo.xml`, SHA-256 `9c5426fa71ff68cc7e40e19e02b5992daf91da5754ef643d2db2f89bd70bb122` |
| `Book 2.xlsx` | `formula`, no target risk family |
| `Book 3.xlsx` | `formula`, no target risk family |
| `ErrorsAndWarnings.xlsx` | `formula`, no target risk family |
| `allelements.xlsm` | `metadata-only`, `preservedMacro`, not a signed/unknown replacement |
| `userformvba.xlsm` | `metadata-only`, `preservedMacro`, not a signed/unknown replacement |

The external candidate is useful but not release-ready:

- it is not checked into Ascend;
- it only replaces the unknown-part shape, not signed-workbook evidence;
- the GitHub repository API reports `license: null`, so vendoring needs product/release approval;
- keeping the current generated unknown-part fixture remains more reproducible until fixture acquisition policy is explicit.

## Confidence

Medium-high that a real public unknown-part candidate exists. Medium-low that it can be vendored, because license/provenance policy is unresolved. High that the signed-workbook fixture gap remains open.

## Fold-in decision

Promote to topic synthesis and owner handoff only. Do not vendor the workbook, do not change `safe-open-proof`, and do not mark `public-edge-fixtures` satisfied.

## Next question

Should product/release approve acquisition rules for external workbook fixtures with unclear repository-level licensing, or should the safe-open proof continue using disclosed generated topology fixtures until a cleanly licensed public signed/unknown fixture appears?
