# Range-Backed Journal Validation

## Question

Can the in-flight journal validation sweep classify invalid range-backed operations as structured `value-unsupported` journal evidence, and can the package graph keep content-type parsing coverage for XML attribute variants?

## Hypothesis

Yes. Invalid ranges for range-backed operations should fail closed in journal construction with stable surface/reason fields instead of leaking parser-specific errors or missing rollback evidence. Content-type parsing should keep handling normal OPC `Default` and `Override` attributes with whitespace, quote, and entity variants because package feature classification depends on it.

## External sources checked

- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- ECMA-376 Open XML standards page: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- OPC content types stream markup reference: https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Content_topic_ID0ESTAI.html

## Why this matters to Ascend

The auditable mutation claim depends on structured, explainable failure evidence for invalid operations, not just successful writes. The safe-open/package-action claims also depend on package graph classification, which starts with `[Content_Types].xml`.

## Probe/implementation

- Finished the in-flight SDK journal change by formatting and validating `packages/sdk/src/journal-exactness.test.ts`.
- `packages/sdk/src/journal.ts` now routes invalid ranges for `fillFormula`, `clearRange`, `setNumberFormat`, `setStyle`, `sortRange`, `copyRange`, `moveRange`, `mergeCells`, and `unmergeCells` into `UNSUPPORTED_VALUE` journal issues.
- Added package graph coverage for content-type attributes with XML whitespace, single quotes, entity decoding, and similarly named attributes that should not be matched.

## Results

- Targeted validation passed:
  - `bun test packages/io-xlsx/src/package-graph.test.ts packages/sdk/src/journal-exactness.test.ts`
  - `bunx biome check packages/io-xlsx/src/reader/content-types.ts packages/io-xlsx/src/package-graph.test.ts packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts`
- The new journal test covers 9 invalid range-backed operation shapes and asserts `code=UNSUPPORTED_VALUE`, stable surfaces, `reason=value-unsupported`, and matrix allowance.

## Confidence

High for the narrow journal classification and content-type regression coverage. Medium for broader journal law claims; this is another boundary hardening step, not a full inverse-law proof.

## Fold-in decision

Promote to correctness loop. This is a small in-flight validation fix under the existing auditable mutation claim. Do not promote a new surface or new release claim wording.

## Next question

Should the next correctness proof focus on remaining invalid operation classes, or should claim stewardship return to owner gates now that range-backed journals are classified?
