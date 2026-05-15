# Legacy Array Binding Member Proof

Date: 2026-05-15

## Question

Can Ascend's verifier reject a legacy array formula range when an occupied cell inside the declared range is detached from the array binding?

## Hypothesis

If the verifier treats each occupied cell inside a legacy array range as part of the binding proof, then copied or corrupted array ranges cannot silently retain stale scalar values inside an array-owned region.

## External sources checked

- Open XML SDK `CellFormula` reference, including formula type and reference attributes: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula
- Microsoft Excel dynamic array behavior background: https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531
- LSP 3.17 rename/prepareRename refusal model for the formula-intelligence boundary: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/

## Why this matters to Ascend

The auditable package-part mutation claim depends on being honest about workbook semantics after mutation. Array-formula ranges are a correctness boundary: if an occupied member cell is detached from its range binding, a plan/commit proof should fail closed instead of treating the workbook as structurally valid.

## Probe/implementation

- Added verifier range-member validation for legacy array formula bindings.
- A range anchor with `formulaInfo.kind === 'array'` now scans occupied cells inside the range.
- Empty cells are allowed, but occupied cells with no matching array binding produce `legacy-array-range-member-mismatch`.
- The helper compares sheet-qualified binding ranges using existing parsed range semantics.

## Results

Proof commands run:

```bash
bun test packages/verify/src/verify.test.ts -t "detects occupied cells detached inside a legacy array range"
bun test packages/sdk/src/agent-workflow.test.ts -t "quality moat matrix proves release-critical formula trust paths"
```

Observed evidence:

- The focused verifier test passed with issue kind `legacy-array-range-member-mismatch` and refs `Sheet1!A1`, `Sheet1!A2`.
- The SDK quality moat test passed with the new detached legacy-array case included in the release-critical formula trust matrix.

## Confidence

High for the generated in-memory corruption case. Medium for real-world breadth until a public XLSX fixture with legacy array range drift is found or generated through Excel/LibreOffice and checked through reopen behavior.

## Fold-in decision

Promote to correctness loop and fold in as a tiny verifier bug fix. This is correctness hygiene under the auditable mutation claim, not a new formula-intelligence surface and not a rename implementation.

## Next question

Can the formula trust moat use a real public workbook or generated fixture round-tripped through a spreadsheet application to prove the same legacy-array detached-member rejection after XLSX import?
