# Post Write Reopen Package Expectations

## Question

Can prepared post-write verification stay faster for simple formula-free edits while treating generated copy-sheet table parts as expected package graph changes?

## Hypothesis

Yes. For formula-free workbooks, a formula-mode post-write reopen only needs formula/binding evidence and sheet dimension hints; it should not hydrate every scalar value. Separately, a copied table that generates a new `xl/tables/tableN.xml` part should be reported as an expected package content-type override, not as an unresolved audit failure.

## External sources checked

- Microsoft OPC overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- OOXML content types stream reference: https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Content_topic_ID0ESTAI.html
- OOXML table definition part reference: https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Table_topic_ID0ERQCM.html
- Microsoft Open XML SDK `TableParts` reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tableparts?view=openxml-2.20.0

## Why this matters to Ascend

Prepared commits should be both trustworthy and cheap. The post-write phase needs to distinguish real package drift from expected generated parts, and it should avoid rehydrating scalar cells when a formula-free workbook only needs structural and formula absence evidence.

## Probe/implementation

Implemented:

- `formulaModeHydrateValues` read/open option, normalized through SDK session load options.
- formula-only no-value fast path that preserves dimensions and density hints when sheet data contains no formulas.
- simple `setCells` post-write reopen now disables formula-mode value hydration only when the source workbook has no formula cells.
- expected post-write package graph changes now include generated copied table part paths and classify their content-type override issues as expected warnings.

## Results

Validation:

- `bun test packages/io-xlsx/src/reader/reader.test.ts -t "formula-only"`
- `bun test packages/sdk/src/agent-workflow.test.ts -t "prepared copySheet commits reopen workbook-unique table identities"`
- `bun test packages/sdk/src/agent-workflow.test.ts -t "post-write"`
- `bunx biome check packages/io-xlsx/src/reader/index.ts packages/io-xlsx/src/reader/sheet.ts packages/io-xlsx/src/reader/reader.test.ts packages/sdk/src/agent-workflow.ts packages/sdk/src/agent-workflow.test.ts packages/sdk/src/load.ts packages/sdk/src/session.ts`
- `bunx tsc --build`
- `bun run test:changed`

## Confidence

Medium. The targeted tests cover formula-free no-value hydration and copied-table post-write expectations, and the full changed-test gate passed. More confidence would come from a timed post-write profile on large formula-free workbooks.

## Fold-in decision

Promote to correctness and performance loops as post-write verification hygiene. Do not promote a release performance claim yet.

## Next question

Can a timed public/generated formula-free prepared-commit profile quantify the benefit without turning local timing into release threshold wording?
