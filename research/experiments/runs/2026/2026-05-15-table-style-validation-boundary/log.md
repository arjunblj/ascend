# Table Style Validation Boundary

## Question

Should `setTableStyle` reject invalid public style metadata before mutating table style state, and should journal exactness classify the same rejected values?

## Hypothesis

Yes. Table style names should be non-empty strings or `null`, and table style display flags should be booleans. Rejecting invalid values before mutation keeps table metadata auditable and avoids truthy string/number surprises.

## External sources checked

- OOXML `tableStyleInfo` element reference: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_tableStyleInfo_topic_ID0EVRU5.html
- Microsoft Open XML SDK `TableStyleInfo`: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyleinfo
- Microsoft Excel table style options: https://support.microsoft.com/en-us/office/format-an-excel-table-6789619f-c889-495c-99c2-2f971c0e2370

## Why this matters to Ascend

The auditable package-part mutation claim depends on public operations staying symbolic and explainable. Table style metadata is user-visible workbook state; invalid style names or non-boolean stripe/first/last-column flags should not mutate a workbook and later become hard-to-explain journal entries.

## Probe/implementation

An in-flight fix added:

- engine validation for invalid `setTableStyle.styleName`, `showFirstColumn`, `showLastColumn`, `showRowStripes`, and `showColumnStripes`.
- engine regression coverage proving invalid inputs reject without changing `tableStyleInfo`.
- SDK journal exactness classification for the same invalid public inputs as `UNSUPPORTED_VALUE`.

## Results

Focused validation:

- `bun test packages/engine/src/operations.test.ts -t "setTableStyle rejects invalid"`
- `bun test packages/sdk/src/journal-exactness.test.ts -t "classifies missing required metadata updates as unsupported values"`
- `bunx biome check packages/engine/src/operations/table-ops.ts packages/engine/src/operations.test.ts packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts`
- `bunx tsc --build`

Full-suite validation was also run while the diff was present:

- `bun run test:changed` passed with 5199 pass, 1 skip, 0 fail.

## Confidence

High. The fix is narrow and mirrors the same engine/journal symmetry used for recent metadata validators.

## Fold-in decision

Promote to correctness loop as auditable mutation hygiene. Do not promote new table-formatting claims or release wording.

## Next question

If the remaining RC gate and IO reader threshold changes are coherent, validate them as release/performance work; otherwise keep the next cycle research-only until the worktree clears.
