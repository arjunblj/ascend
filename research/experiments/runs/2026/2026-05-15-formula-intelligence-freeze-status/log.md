# Formula Intelligence Freeze Status

## Question

Can formula intelligence be frozen in the owner handoff with a final rejection-first status row, so future loops do not reopen rename implementation during this claim-steward block?

## Hypothesis

Yes. The formula-assist proof supports formula-local primitives and refusal classification, not edit-producing rename. The owner handoff should preserve that boundary.

## External sources checked

- LSP 3.17 `prepareRename`: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Excel LET function: https://support.microsoft.com/en-gb/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
- Excel structured references: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- HyperFormula named expressions: https://hyperformula.handsontable.com/guide/named-expressions.html

## Why this matters to Ascend

Formula intelligence is valuable, but unsafe rename would cut against the North Star. Workbook-context names, table columns, external refs, sheet refs, and ranges require ownership and operation-specific rewrite planning that formula assist does not provide.

## Probe/implementation

Ran:

```bash
bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json
```

Updated `research/experiments/syntheses/2026-05-owner-handoff.md` with a formula freeze status row.

## Results

| Field | Value |
| --- | --- |
| Public formulas discovered | 1685 |
| Sampled formulas | 1685 |
| Static edge cases | 10 |
| Parse OK count | 1695 |
| Diagnostic formula count | 0 |
| Reference spans | 2322 |
| Binding roles | 25 |
| LET-local prepare-rename OK targets | 3 |
| `no-symbol-at-cursor` refusals | 285 |
| `workbook-context-required` refusals | 4 |
| `reference-target-not-renameable` refusals | 1403 |
| Proof passed | true |

Boundary: formula assist proves formula-local assist latency and rejection-first `prepareRename` classification. It does not apply workbook edits, resolve workbook names, or claim safe cross-workbook/table rename.

## Confidence

High. The proof backs refusal-first language-service primitives and explicitly rejects workbook-context rename targets.

## Fold-in decision

Promote to topic synthesis only. Do not implement rename. Keep formula intelligence out of the top implementation handoffs until workbook-context symbol ownership and operation-owned edits exist.

## Next question

Can the next research cycle safely move away from release-claim stewardship and return to a single performance unknown, or should owner-loop approval remain the blocker before new experiments?
