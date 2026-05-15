# Formula Prepare-Rename Rejection Matrix

## Question

Can the formula language-service claim be tightened with rejection-first evidence for defined names, table symbols, sheet refs, external refs, 3D refs, spills, literals, functions, and parse failures without implementing rename?

## Hypothesis

Yes. `formulaPrepareRename()` already exposes a guard result and should return `ok: true` only for formula-local `LET` bindings. A local probe should catch any target that is incorrectly treated as a renameable or workbook-context name when it is really a cell/range/sheet/external reference.

## External sources checked

- LSP 3.17 `prepareRename` can produce no result for invalid rename targets: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Microsoft names in formulas define workbook and worksheet scope and local-name precedence: https://support.microsoft.com/en-us/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1
- Microsoft structured references document table and column names and workbook-wide updates when tables or columns are renamed: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Microsoft formula overview documents worksheet references and quoted sheet names: https://support.microsoft.com/en-us/office/formulas-and-functions-294d9486-b332-48ed-b489-abe7d0f9eda9

## Why this matters to Ascend

The claim is "formula language-service primitives," not safe rename. If a guard misclassifies sheet or external workbook qualifiers as unresolved names, a future client could overread the evidence and attempt a symbol rename where only a workbook operation should be allowed.

## Probe/implementation

- Probed `formulaPrepareRename()` on:
  - defined names;
  - table names and columns;
  - structured-reference item selectors;
  - cell, range, sheet, 3D, spill, and external references;
  - function names, literals, and parse-failure punctuation.
- Found a bug: cursors on sheet/external/3D reference qualifiers were classified as unresolved names and returned `workbook-context-required`.
- Folded in a scoped guard fix:
  - non-structured references are rejected before binding-role lookup with `reference-target-not-renameable`;
  - structured references still use table-name/table-column roles when available;
  - structured item selectors without a role now reject with `workbook-context-required`.
- Added an SDK rejection matrix test.
- Did not add any edit-producing rename operation.

## Results

Validation passed:

```bash
bun test packages/sdk/src/formula-edit.test.ts -t "prepare|binding roles|formula IDE"
bun test apps/cli/src/cli.test.ts -t "formula assist"
bun test apps/api/src/server.test.ts -t "formula"
bun test apps/mcp/src/index.test.ts -t "formula_assist|formula assist"
bunx biome check packages/sdk/src/formula-edit.ts packages/sdk/src/formula-edit.test.ts
bunx tsc --build
bun run test:changed
```

Validation note: the broader `test:changed` gate exposed stale saved-source journal exactness expectations and a type-narrowing issue in journal package-state refs. Those were fixed separately so this formula fold-in lands with the full suite green.

Rejection evidence now matches the claim-board contract:

| Cursor target | Result |
| --- | --- |
| resolved `LET` declaration/use | `ok: true`, formula-local occurrence ranges only |
| defined name | `ok: false`, `workbook-context-required` |
| table name or table column | `ok: false`, `workbook-context-required` |
| structured item selector | `ok: false`, `workbook-context-required` |
| cell/range/sheet/3D/spill/external ref | `ok: false`, `reference-target-not-renameable` |
| function/literal/punctuation/parse failure | `ok: false`, `no-symbol-at-cursor` |

## Confidence

High for the rejection matrix covered by SDK tests. Medium for the broader formula intelligence claim until workbook-context defined-name/table ownership and cross-surface proof packaging are stronger.

## Fold-in decision

Promote as a tiny correctness/product fold-in to formula language-service primitives. Keep safe rename and edit-producing rename on the "do not promote" list.

## Next question

Can retained viewport patch history receive a current proof rerun and cross-surface boundary check without touching writer code?
