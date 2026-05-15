# Package Action DocProps Public Fixture

## Question

Can the auditable package-part mutation proof replace the generated `docprops-passthrough` package with a checked-in public XLSX fixture while preserving the same passthrough evidence?

## Hypothesis

Yes. Core, extended, and custom document properties are ordinary OPC package parts, so a real public workbook that already carries `docProps/*` should prove passthrough better than a synthetic package when Ascend edits only worksheet cell content.

## External sources checked

- Microsoft OPC overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- OOXML core file properties part reference: https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Core_topic_ID0ED3CO.html
- OOXML OPC core properties reference: https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Core_topic_ID0EBYDI.html
- Microsoft Open XML SDK `CoreFilePropertiesPart`: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.corefilepropertiespart?view=openxml-3.0.1
- Apache POI OPC relationship constants: https://poi.apache.org/apidocs/dev/org/apache/poi/openxml4j/opc/PackageRelationshipTypes.html

## Why this matters to Ascend

The top package-action release blocker is fixture disclosure or replacement. Replacing generated edge packages with checked-in public fixtures makes the "auditable package-part mutation" claim more credible without adding new SDK, CLI, API, or MCP surfaces.

## Probe/implementation

Inspected public XLSX fixtures for `docProps/core.xml` and `docProps/custom.xml`, then ran a local package-action probe against candidate workbooks. `fixtures/xlsx/calamine/date_1904.xlsx` produced byte-equal passthrough actions for `docProps/core.xml`, `docProps/app.xml`, and `docProps/custom.xml` while regenerating only changed workbook content.

Folded that candidate into `fixtures/benchmarks/package-action-proof.ts` as the `docprops-passthrough` case and updated release-proof index wording/tests so generated edge-package acceptance now refers only to signature and unknown structural packages.

## Results

- `docprops-passthrough` is now `sourceKind: public-fixture` with fixture `fixtures/xlsx/calamine/date_1904.xlsx`.
- Package-action compact proof now reports source counts `public-fixture=4`, `generated-workbook=2`, `generated-edge-package=2`.
- Combined package action counts remain all-action coverage: `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1`.
- The release proof index remains fail-closed: `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, and `missingRequirementCount=9`.
- New package-action stable-shape digest: `b4cf0755f2fd40ff577ab87b6600ee14a58bbd69608796c903759e7e8b4d29e8`.

Validation:

```bash
bun test fixtures/benchmarks/package-action-proof.test.ts
bun test fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

## Confidence

High for replacing the docProps passthrough proof with a public fixture. Medium for broader fixture-policy progress: signature invalidation and unknown-part error still require generated structural packages or owner-approved public replacements.

## Fold-in decision

Promote to product/release proof packaging. This is a scoped benchmark proof fold-in, not a production mutation surface. Keep `edge-fixture-policy` missing because generated signature and unknown-part cases remain.

## Next question

Can the package-action proof replace `signature-invalidation-drop` or `unknown-part-error` with license-clear public binary fixtures, or should product explicitly accept disclosed generated structural packages for those two remaining cases?
