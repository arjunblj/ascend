# Shared Formula Journal Coverage

## Question

Does a `moveRange` that rewrites formulas outside the moved range journal every affected shared-formula member, or can the audit under-report the modified formula surface?

## Hypothesis

Formula rewrite journaling was collecting plain cell preimages for rewritten cells. For shared formula members, that can miss binding-metadata loss issues that `cellEditPreimages` already reports for direct edits.

## External sources checked

- Open XML SDK `CellFormula` docs for SpreadsheetML formulas and shared-formula attributes: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula
- Microsoft Open XML SDK shared-string/formula cell structure overview: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas
- ClosedXML formula calculation docs, which note that shared formulas are an XLSX memory optimization: https://docs.closedxml.io/en/latest/concepts/formula-calculation.html

## Why this matters to Ascend

The auditable package-part mutation claim depends on honest rollback-journal evidence. If a structural edit rewrites multiple shared-formula members but the journal only names one member, agents can under-estimate the affected formula surface before approving or auditing a write.

## Probe/implementation

Inspected `packages/sdk/src/journal.ts` around `moveRangeFormulaSurfaceRestoration`. The function already discovers rewritten formula cells by comparing formula text before and after `rewriteWorkbookFormulasForMove`, but it converted those refs through `cellPreimages`.

Fold-in:

- Switched the rewritten formula cell collection to `cellEditPreimages`, matching direct formula-edit journaling.
- Added a focused journal exactness regression for moving `B1:B2` to `C1` while formulas in shared-formula group `A1:A2` rewrite to `C1*2` and `C2*2`.
- Strengthened the agent workflow quality moat to expect both `Data!B1` and `Data!B2` shared-formula issues for the existing move-range case.

## Results

Validation:

```bash
bun test packages/sdk/src/journal-exactness.test.ts -t "moveRange formula rewrites journal every shared formula member"
bun test packages/sdk/src/agent-workflow.test.ts -t "quality moat matrix proves release-critical formula trust paths"
bunx biome check packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts packages/sdk/src/agent-workflow.test.ts
git diff --check -- packages/sdk/src/journal.ts packages/sdk/src/journal-exactness.test.ts packages/sdk/src/agent-workflow.test.ts
```

All passed.

## Confidence

High for this edge case: the regression proves both affected shared-formula members now appear in journal issue refs. Medium for broader structural formula rewrite coverage because this does not exhaust every shared-formula import shape.

## Fold-in decision

Promote to correctness loop and auditable package-part mutation proof hygiene. This is a small production correctness fix, not a new product surface.

## Next question

Can practical latency target selection report enough tail/variance evidence to avoid optimizing noisy medians without turning local timing into release claims?
