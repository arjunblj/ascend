# Conditional Format Order Journal Boundary

Date: 2026-05-15

## Question

Can conditional-format replacement/deletion journals honestly distinguish exact inverse restoration from cases where public operations can restore metadata content but not saved package order?

## Hypothesis

Yes. Conditional format metadata is order-sensitive at the saved package level. Ascend should keep tail replacements exact, but classify non-tail replacement/deletion as `LOSSY_INVERSE` with `surface: "conditional-formats"` and `reason: "metadata-order"` when public inverse operations cannot restore the original ordering.

## External Sources Checked

- Microsoft Open XML conditional formatting documentation defines `<conditionalFormatting/>` and `<cfRule/>` as SpreadsheetML package elements, and notes conditional formatting rules applied to cell/range collections: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting
- The Open XML SDK `ConditionalFormatting` class maps directly to the serialized `x:conditionalFormatting` element: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.conditionalformatting?view=openxml-3.0.1

## Why This Matters To Ascend

This strengthens the auditable mutation claim by reducing overclaim. A mutation journal that says "exact" must restore workbook evidence and saved package bytes for the cases covered by that test. If public operations can only append restored conditional-format metadata, Ascend should expose that boundary instead of silently producing different package order.

## Probe/Implementation

Folded in a scoped SDK journal change:

- `journalSetConditionalFormat` now adds a metadata-order lossy issue when a replacement preimage is not a suffix of the current conditional-format list.
- `journalDeleteConditionalFormat` does the same for range/priority/rule-index deletes, while preserving the existing whole-sheet delete boundary.
- Added direct tests for conditional-format replacement and delete order classification.
- Expanded the representative exact saved-byte test with additional exact operation families, but kept conditional-format replacement exact only for the tail case that public inverse operations can restore.

## Results

Validation passed:

```bash
bun test packages/sdk/src/journal-exactness.test.ts
bun test packages/sdk/src/interactive-contract.test.ts -t "journal inverse ops restore conditional format replacements|copyRange format journals restore standard conditional formats exactly|moveRange format journals restore standard conditional formats exactly"
bunx biome check packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts
bunx tsc --build packages/sdk/tsconfig.json
```

Observed proof:

- Non-tail replacement of `Sheet1!A1:A2` while another conditional format follows now yields:
  - `code: "LOSSY_INVERSE"`
  - `surface: "conditional-formats"`
  - `reason: "metadata-order"`
  - `refs: ["Sheet1!A1:A2"]`
- Non-tail deletion of `Sheet1!A1:A2` while another conditional format follows yields the same metadata-order issue.
- Tail replacement remains in the representative exact saved-byte test and passes byte restoration.

## Confidence

High for the conditional-format order boundary. Medium for all conditional-format variants, because x14 conditional formats and more complex priority interactions remain covered by separate tests and should not be collapsed into this classic conditional-format result.

## Fold-In Decision

Promote to correctness loop and commit. This is a productive "do not overclaim exactness" fold-in for auditable mutation, not a new product surface.

## Next Question

Should the next journal-law proof add property or matrix coverage for non-tail conditional-format deletes and replacement priority reassignment, or should the next loop return to top-claim proof packaging?
