# Case-Insensitive Copy-Sheet Binding Proof

## Question

Can the SDK quality moat prove that `copySheet` retargets imported shared-formula binding metadata even when sheet-qualified binding refs use different casing?

## Hypothesis

Yes. A generated workbook with `data!A1` and `DATA!A1:A2` shared-formula metadata should copy from sheet `Data` to `Copy`, reopen cleanly, and contain only `Copy!` binding refs.

## External sources checked

- Microsoft Learn, working with formulas in SpreadsheetML: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- OOXML shared formulas primer: https://c-rex.net/samples/ooxml/e1/Part3/OOXML_P3_Primer_Shared_topic_ID0EVFGK.html

## Why this matters to Ascend

Auditable mutation claims depend on saved workbooks reopening with trustworthy formula metadata, not just normalized formula text. Imported XLSX files can carry sheet-qualified refs with casing drift; copy-sheet mutation must not preserve stale source-sheet binding refs in the copied worksheet.

## Probe/implementation

Added a case to `packages/sdk/src/agent-workflow.test.ts` quality moat:

- create sheet `Data`;
- seed a shared-formula group with lowercase and uppercase sheet-qualified binding refs;
- save, prepare `copySheet` to `Copy`, commit, reopen;
- assert no formula-binding integrity issues;
- assert copied formulas normalize to `Copy!B1*2` and `Copy!B2*2`;
- assert copied binding metadata uses `masterRef=A1` and `ref=Copy!A1:A2`;
- assert copied binding JSON does not contain `data!` case-insensitively.

## Results

Validation passed:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "quality moat matrix proves release-critical formula trust paths"
bun run test:changed
```

The focused quality-moat test passed with 184 assertions. The changed-test gate passed with 5135 tests, 1 skip, 0 failures.

## Confidence

High for this regression scope. This is test-only correctness evidence and does not add a new SDK/API/MCP surface.

## Fold-in decision

Promote to correctness-loop evidence under auditable mutation and formula-binding trust. Do not promote as a product claim.

## Next question

Should future formula-binding trust moats use a shared helper to generate casing-drift bindings across copy, move, and structural edit operations?
