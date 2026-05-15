# CopyRange Style Binding Regression

## Question

Do style-only copy modes preserve target formula binding metadata when the copied source only contributes formatting?

## Hypothesis

Yes. `formats` and `styles` paste modes should update style information without detaching non-spill formula metadata for shared formulas or data-table formula groups.

## External sources checked

- [Microsoft Support: move or copy cells and cell contents](https://support.microsoft.com/en-us/office/move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) frames copy behavior as carrying selected cell contents and formatting depending on paste mode.
- [Microsoft Support: move or copy a formula](https://support.microsoft.com/en-us/office/move-or-copy-a-formula-1f5cf825-9b07-41b1-8719-bf88b07450c6) documents that formula semantics and references need separate handling from formatting changes.
- [OOXML shared formulas primer](https://c-rex.net/samples/ooxml/e1/Part3/OOXML_P3_Primer_Shared_topic_ID0EVFGK.html) explains shared formula metadata and why member cells can rely on a primary formula.

## Why this matters to Ascend

Style-only edits are common agent and human workbook mutations. If a formatting-only copy strips shared-formula or data-table metadata, Ascend would convert a cosmetic edit into a formula semantics mutation, weakening preservation-first claims.

## Probe/implementation

Finished the in-flight regression coverage by adding a focused engine test for `copyRange` with `formats` and `styles` paste modes. The test covers:

- shared formula member metadata after style-only copy,
- data-table formula metadata after style-only copy,
- unchanged recalculation requirement for both modes,
- affected-cell reporting limited to the formatted target.

## Results

- `bun test packages/engine/src/operations.test.ts -t "copyRange style-only paste preserves non-spill formula binding metadata"` passed: 1 test.

## Confidence

Medium. This proves two important non-spill binding cases for style-only copy. It does not prove dynamic-array spill semantics or every style-bearing metadata surface.

## Fold-in decision

Promote to correctness loop as regression coverage only. No production surface and no formula rename work.

## Next question

Should formula-binding regressions be grouped into a small generated operation matrix so style-only, value, formula, and all-mode paste behavior can be audited together?
