# Sparkline Group Validation Boundary

## Question

Should `setSparklineGroup` reject invalid public x14 sparkline metadata before mutating the workbook model?

## Hypothesis

Yes. Sparkline groups are x14 extension metadata with range formulas and boolean display flags. Invalid ranges, empty types, or non-boolean flags should fail before mutation so auditable workbook edits never create metadata the writer cannot represent honestly.

## External sources checked

- Microsoft Open XML SDK `SparklineGroup`, which documents `x14:sparklineGroup`, range formula children, sparkline type, marker/high/low/first/last/negative/displayXAxis boolean attributes, and Office 2010 availability: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.sparklinegroup?view=openxml-3.0.1
- Microsoft Open Packaging Conventions overview for package parts and relationship boundaries around extension metadata: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

Sparkline metadata is not a headline claim, but it sits under the auditable mutation claim. If public operations can inject invalid x14 ranges or flag values, follow-on package evidence becomes harder to trust and journal exactness boundaries become fuzzier.

## Probe/implementation

- Inspected the in-flight patch in `packages/engine/src/operations/visual-ops.ts`.
- Added validation for `range` and `locationRange` through the existing A1 range parser.
- Added validation for non-empty sparkline `type`.
- Added validation that public boolean fields remain booleans.
- Added focused engine tests proving invalid values reject before the existing sparkline group is mutated.

## Results

- Targeted validation passed:
  - `bun test packages/engine/src/operations.test.ts -t "setSparklineGroup"`
  - `bunx biome check packages/engine/src/operations/visual-ops.ts packages/engine/src/operations.test.ts`
  - `bunx tsc --build`
  - `bun run test:changed`
- Fold-in scope is correctness-only. This does not add a new surface, claim, or benchmark threshold.

## Confidence

High for the validation boundary. The code reuses existing range parsing and only guards fields already accepted by the public operation.

## Fold-in decision

Promote to correctness loop and commit as small production validation hardening under auditable mutation.

## Next question

Should the package-action `unknown-part-error` gate get an external-candidate owner-review record, or should it stay generated-only because the ExcelForge workbook is not proven as a fail-closed mutation case?
