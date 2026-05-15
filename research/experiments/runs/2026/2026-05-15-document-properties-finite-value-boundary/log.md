# Document Properties Finite Value Boundary

## Question

Should document-property metadata reject `NaN` and infinite numeric values before workbook mutation?

## Hypothesis

Yes. Open XML document-property values are serialized as typed XML property values. JavaScript non-finite numbers cannot be represented as honest numeric workbook metadata and should fail before mutation with auditable journal issue classification.

## External sources checked

- Microsoft Open XML custom-property guidance, which describes custom properties as typed values such as text, yes/no, integer, and double values: https://learn.microsoft.com/en-us/office/open-xml/word/how-to-set-a-custom-property-in-a-word-processing-document
- Microsoft Open XML SDK `CustomDocumentProperty`, which maps custom properties to typed `vt:*` value elements such as booleans, strings, and numeric value types: https://learn.microsoft.com/ru-ru/dotnet/api/documentformat.openxml.customproperties.customdocumentproperty?view=openxml-3.0.1
- Microsoft Open XML SDK `ExtendedProperties.Properties`, which documents the extended application properties part used for Office document metadata: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.extendedproperties.properties?view=openxml-3.0.1

## Why this matters to Ascend

Ascend's auditable mutation claim depends on keeping workbook metadata representable and explainable. If `NaN` or `Infinity` can enter document properties, the writer must either invent serialization behavior or emit invalid metadata, both of which weaken preservation-first package mutation.

## Probe/implementation

- Inspected the active dirty patch in `packages/engine/src/operations/workbook-ops.ts`.
- Kept the implementation narrow: app document properties and scalar-array entries now use a shared finite-value guard, and custom property numeric values require `Number.isFinite`.
- Added engine tests proving invalid scalar metadata rejects before mutation for core, app, app scalar-array, custom-value, and custom-name cases.
- Added SDK journal exactness cases proving the same invalid inputs classify as workbook-metadata `UNSUPPORTED_VALUE` issues.

## Results

- Targeted validation passed:
  - `bun test packages/engine/src/workbook-ops.test.ts packages/sdk/src/journal-exactness.test.ts`
  - `bunx biome check packages/engine/src/operations/workbook-ops.ts packages/engine/src/workbook-ops.test.ts packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts`
- Fold-in scope is correctness-only. This does not add a product surface or change benchmark thresholds.

## Confidence

High. The guard matches the existing SDK journal classifier shape and prevents non-representable numeric metadata from entering the workbook model.

## Fold-in decision

Promote to correctness loop and commit as a tiny production validation fix under the auditable mutation claim.

## Next question

Can the safe-open external unknown-part workbook candidate become machine-readable owner-review evidence in the release proof index without satisfying the public-edge fixture gate?
