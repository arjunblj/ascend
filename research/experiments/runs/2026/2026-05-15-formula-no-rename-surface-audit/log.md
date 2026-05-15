# Formula no-rename surface audit

## Question

Do Ascend's current formula-intelligence surfaces stay inside the rejection-first boundary, or did the loop accidentally promote edit-producing rename?

## Hypothesis

The surfaces remain bounded. SDK/CLI/API/MCP expose `renameTarget` metadata and formula-local LET prepare evidence, but no workbook edit, operation planner, `WorkspaceEdit`, or cross-workbook rename surface is exposed by formula assist.

## External sources checked

- LSP 3.17 `prepareRename` separates rename preparation from edit production and allows refusal: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft LET describes formula-local names defined inside one formula: https://support.microsoft.com/en-au/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
- Microsoft names-in-formulas documentation describes workbook- and worksheet-scoped names: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft structured references documentation describes table and column symbols in formulas: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e

## Why this matters to Ascend

Formula intelligence is useful, but unsafe rename would undermine trustworthy mutation planning. Ascend can claim language-service primitives only if the surface refuses workbook-context symbols and references until a future operation-owned planner proves the full edit set.

## Probe/implementation

Searched formula-assist and rename-adjacent code paths:

```bash
rg -n "prepareRename|renameTarget|formulaAssist|formula[-_ ]assist|rename target|occurrenceRanges|apply.*rename|WorkspaceEdit|TextEdit" packages apps fixtures research/experiments/syntheses -g '*.ts' -g '*.md'
rg -n "prepareRename|renameTarget|occurrenceRanges|workbook-context-required|reference-target-not-renameable" packages/sdk packages/formulas apps/cli apps/api apps/mcp fixtures/benchmarks -g '*.ts'
```

Inspected:

- `packages/sdk/src/formula-edit.ts`
- `packages/sdk/src/formula-edit.test.ts`
- `apps/cli/src/commands/formula.ts`
- `apps/api/src/server.ts`
- `apps/mcp/src/index.ts`
- `fixtures/benchmarks/formula-assist-proof.ts`

Validation:

```bash
bun test packages/sdk/src/formula-edit.test.ts fixtures/benchmarks/formula-assist-proof.test.ts
bun test apps/cli/src/cli.test.ts -t "formula assist"
bun test apps/api/src/server.test.ts -t "formula-assist"
bun test apps/mcp/src/index.test.ts -t "formula_assist"
```

## Results

- `formulaPrepareRename` returns `ok: true` only for formula-local LET binding evidence and explicitly says callers must still apply edits.
- Workbook names, table names, table columns, structured selectors, defined-name-like tokens, cell/range refs, sheet refs, 3D refs, spill refs, and external workbook refs return refusal reasons.
- CLI/API/MCP forward `formulaAssist` results; they do not expose a formula rename command, mutation operation, `WorkspaceEdit`, or workbook-wide edit plan.
- Focused validation passed:
  - SDK/formula proof: 22 tests.
  - CLI formula assist: 1 test.
  - API formula-assist: 1 test.
  - MCP formula_assist: 1 test.

## Confidence

High for current surfaces: the scan found formula-assist metadata and existing workbook operations such as sheet/table rename, but no formula-assist edit-producing rename surface. Medium for future drift risk because the SDK exports `formulaPrepareRename`, so release wording should continue saying "guard" or "prepare evidence," not "rename."

## Fold-in decision

Promote to topic synthesis and keep formula intelligence out of implementation handoff. Do not implement rename in this block.

## Next question

Should the release claim board add a recurring proof-owner gate that rejects formula rename wording unless a future workbook-context resolver and operation-owned edit planner are both present?
