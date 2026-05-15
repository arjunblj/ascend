# Package Action Generated Edge Streaming Probe

## Question

Should package-action proof expand streaming coverage to generated signature and unknown/error edge cases, or should generated edge streaming remain outside release wording?

## Hypothesis

If generated edge streaming produces the same expected package actions as the standard writer proof, then the performance loop could expand the streaming matrix. If unknown-part streaming does not preserve the fail-closed `error` action, then full streaming parity must stay blocked and the release claim should remain representative-only.

## External sources checked

- OPC package parts and relationships overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- OOXML digital signature origin relationship: https://c-rex.net/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Digital_topic_ID0EGZAO.html
- OOXML digital signature parts: https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Digital_topic_ID0EOIAK.html
- OOXML package digital signature constraints: https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Digital_topic_ID0EHROM.html

## Why this matters to Ascend

The package-action claim should not drift from "auditable package-part mutation" into "streaming parity for every edge package." Generated signature and unknown-part workbooks are exactly where overclaiming would be expensive: signatures imply provenance/trust, and unknown parts must fail closed rather than silently losing review semantics.

## Probe/implementation

Used the existing `defaultPackageActionProofCases()` to prepare the generated edge inputs, then ran a throwaway streaming path with:

- `readXlsx`
- `applyOperations`
- `summarizePlannedWrite`
- `writeXlsxStreaming`
- `createPackageActionProof`

Cases probed:

- `signature-invalidation-drop`
- `unknown-part-error`

No production code changed.

## Results

| Case | Probe result | Decision |
| --- | --- | --- |
| `signature-invalidation-drop` | Streaming proof completed with `drop=2`, `regenerate=5`, no issues, and signature drop present. | Interesting but generated-edge only; not enough to close the owner gate. |
| `unknown-part-error` | Streaming proof completed with `passthrough=1`, `regenerate=5`, no issues, and no `error` action for `custom/custom1.xml`. | Do not fold in; it does not preserve the fail-closed standard-writer proof semantics. |

The public streaming matrix from the previous cycle remains useful, but generated edge/error streaming must stay out of release wording.

## Confidence

High that unknown-part streaming is not ready for package-action release proof. Medium that signature streaming could eventually be folded after owner policy, but it should not be promoted while generated edge fixtures remain owner-gated and signature wording must stay below provenance/trust claims.

## Fold-in decision

Archive as a kill/downgrade result. Do not fold generated edge streaming into the proof harness. Keep `streaming-matrix-boundary` missing and forbid full streaming parity, generated edge/error streaming, arbitrary unknown-part preservation, and signed-provenance wording.

## Next question

Should the release matrix downgrade package-action streaming wording to "five public/generated-workbook representative cases" permanently, or should performance explicitly fund unknown-part fail-closed streaming semantics?
