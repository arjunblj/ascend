# Worksheet MC Stripped Metadata Proof

## Question

Should worksheet-level custom sheet views, raw extension lists, and controls be extracted from markup-compatibility-stripped worksheet XML, the same way x14 validations already are?

## Hypothesis

Yes. Metadata extraction should operate on the normalized worksheet XML after markup compatibility expansion/stripping, otherwise preserved custom views, extension payloads, or controls can be missed or mis-scoped in real OOXML producers that use `mc:AlternateContent` or ignorable namespaces.

## External sources checked

- ECMA-376 defines OOXML package and markup compatibility standards for spreadsheet documents: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- Microsoft Open XML docs describe `AlternateContent` as markup compatibility content with fallback/choice behavior: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.alternatecontent
- Microsoft Open XML docs describe extension lists as OOXML extension containers: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.extensionlist

## Why this matters to Ascend

Preservation-first XLSX depends on reading workbook metadata exactly enough to explain what will pass through, regenerate, or be dropped. If some worksheet metadata is extracted from unstripped XML while adjacent x14 metadata is extracted from stripped XML, package-action and safe-open evidence can drift across producers.

## Probe/implementation

Finished the in-flight reader fix in `packages/io-xlsx/src/reader/sheet.ts`:

- `extractCustomSheetViews(strippedXml, sheet)`
- `extractExtLst(strippedXml, sheet)`
- `extractControls(strippedXml, sheet)`

The SDK trust-moat test also moved a shared-formula corruption proof from a synthetic in-memory workbook to the real `fixtures/xlsx/poi/shared_formulas.xlsx` import path, proving imported shared formula corruption stays blocked even with explicit approvals.

Commands run:

```bash
bun test packages/sdk/src/agent-workflow.test.ts
bun test packages/io-xlsx/src
bunx biome check packages/io-xlsx/src/reader/sheet.ts packages/sdk/src/agent-workflow.test.ts
bunx tsc --build
bun run test:changed
```

## Results

- SDK agent workflow tests passed: 74 tests, 496 assertions.
- io-xlsx package tests passed: 389 tests, 2868 assertions.
- Biome passed for the touched reader/test files.
- `bunx tsc --build` passed.
- `bun run test:changed` passed: 4388 tests, 1 skip, 26718 assertions.

## Confidence

High. The reader change is narrow and the targeted plus changed suites cover worksheet metadata, writer preservation, package graph fidelity, and agent write-policy checks.

## Fold-in decision

Promote to correctness loop. This is a scoped preservation-read fix and a real-fixture trust-moat assertion, not a new public surface.

## Next question

Can the release proof index distinguish this kind of correctness hygiene from release-claim readiness, so small preservation fixes do not look like new claim promotion?
