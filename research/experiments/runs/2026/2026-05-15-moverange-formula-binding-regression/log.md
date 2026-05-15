# MoveRange Formula Binding Regression

## Question

Does `moveRange` materialize target-side formula binding groups before overwriting a member when the paste mode is `values` or `formulas`?

## Hypothesis

Yes. The existing engine implementation appears to preserve the group master and detach overwritten members, but the regression proof should cover both formula-preserving and value-only move modes.

## External sources checked

- [Microsoft Support: move or copy cells and cell contents](https://support.microsoft.com/en-us/office/move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) frames move/copy as replacing the target cell contents while carrying formulas depending on paste semantics.
- [Microsoft Support: move or copy a formula](https://support.microsoft.com/en-us/office/move-or-copy-a-formula-1f5cf825-9b07-41b1-8719-bf88b07450c6) documents formula copy/move behavior and the need to preserve intended references.
- [OOXML shared formulas primer](https://c-rex.net/samples/ooxml/e1/Part3/OOXML_P3_Primer_Shared_topic_ID0EVFGK.html) describes shared formula metadata where only the primary formula may need to be loaded and parsed.

## Why this matters to Ascend

Formula binding groups are a core preservation-first edge case. If a range move overwrites a shared or spill formula member without first materializing the group, Ascend can leave stale formula metadata behind or silently corrupt formula intent.

## Probe/implementation

Finished the already-in-flight regression by adding a focused engine test for `moveRange` into a target shared-formula member. The test covers both `values` and `formulas` paste modes and verifies:

- the shared master remains materialized at `A1`,
- the overwritten target cell has no stale `formulaInfo`,
- formula paste preserves the moved formula,
- value paste overwrites with a literal,
- the source cell is cleared.

## Results

- `bun test packages/engine/src/operations.test.ts -t "moveRange formula paste modes materialize target formula bindings before overwriting a member"` passed: 1 test.

## Confidence

Medium. The regression covers shared formula target members for two paste modes. It does not cover every dynamic array, data table, or cross-sheet move variant.

## Fold-in decision

Promote to correctness loop as a regression test only. No new production surface.

## Next question

Should a broader formula-binding move/copy matrix be generated from operation fixtures, or is the existing targeted regression coverage enough until another binding bug appears?
