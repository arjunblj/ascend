# Release Claim Board No-Promotion Refresh

## Question

Can Ascend act as claim steward for the next block by finishing any tiny in-flight fix, then producing a release-claim board that hands off only the top one or two product claims and keeps formula intelligence rejection-first?

## Hypothesis

Yes. The highest-value action is to narrow claims, not to promote another formula or agent surface. Formula intelligence should be framed as primitives plus refusal semantics until cross-surface refusal snapshots and latency evidence exist.

## External sources checked

- LSP 3.17 rename and prepareRename: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft LET function: https://support.microsoft.com/en-gb/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
- Microsoft names in formulas: https://support.microsoft.com/en-gb/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft structured references: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- HyperFormula cell references: https://hyperformula.handsontable.com/docs/guide/cell-references.html

## Why this matters to Ascend

Ascend needs product-shaped claims that can survive release scrutiny. A claim board turns research into accountable proof requirements: fixture, benchmark, surface, validation gate, competitor contrast, and honest boundary. It also prevents formula intelligence from drifting into unsafe rename work before workbook-context proof exists.

## Probe/implementation

- Ran `git status --short --branch`.
- Found an in-flight production fix around dynamic spill anchor matching. Finished validation and committed it separately as `602fb2b9 fix(engine): match canonical spill anchor refs`.
- Inspected `packages/sdk/src/formula-edit.ts`, `packages/sdk/src/formula-edit.test.ts`, and CLI/API/MCP formula-assist tests.
- Reran current formula evidence:
  - `bun test packages/sdk/src/formula-edit.test.ts`
  - `bun test apps/cli/src/cli.test.ts -t "formula assist returns formula IDE help"`
  - `bun test apps/api/src/server.test.ts -t "formula-assist exposes diagnostics"`
  - `bun test apps/mcp/src/index.test.ts -t "ascend.formula_assist exposes formula IDE helpers"`
- Updated `research/experiments/syntheses/2026-05-release-claim-board.md` to make formula stewardship proof-only and remove it from implementation-loop handoff prompts.

## Results

The release board now has the required three columns:

- claim wording allowed today;
- proof still missing;
- owner loop.

Only the top two claims are implementation-loop handoffs:

1. Safe unknown workbook opening.
2. Auditable package-part mutation.

Formula language-service primitives are not promoted. Allowed wording is limited to parse diagnostics, token/reference spans, hover, completions, reference cycling, binding roles, and rejection-first prepare-rename guard behavior. The spec requires refusal for workbook-context names, table names, table columns, structured item selectors, cell/range/sheet/3D/spill/external refs, function names, literals, punctuation, whitespace, and ambiguous parse failures. It permits `ok: true` only for formula-local LET evidence and still does not apply edits.

## Confidence

High that the board now matches the requested stewardship posture. Medium for formula release readiness: SDK rejection coverage is good, but cross-surface `renameTarget` refusal snapshots and latency/corpus evidence remain missing.

## Fold-in decision

Promote to topic synthesis only. Do not implement rename. Do not add new production formula surfaces. Hand off only safe unknown workbook opening and auditable package-part mutation to implementation/proof loops.

## Next question

Can the safe unknown workbook opening proof be made release-ready from existing surfaces with public fixtures and latency evidence, without adding a new product surface?
