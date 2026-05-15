# Package Action Tracked Fixture Scan

## Question

Can the remaining generated package-action edge cases be replaced by tracked public XLSX/XLSM fixtures, or should the owner handoff continue to require explicit generated-fixture approval?

## Hypothesis

The tracked fixture corpus should prove common package features such as docProps, calc chains, custom XML, macros, and chart/drawing sidecars, but likely does not contain digital signature packages or the synthetic unknown-path family needed for the two remaining generated edge cases.

## External sources checked

- Microsoft OPC overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- OOXML unknown relationships: https://ooxml.info/docs/9/9.1/9.1.7/
- OOXML digital signature origin part: https://ooxml.info/docs/15/15.2/15.2.7/
- OOXML digital signatures: https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Digital_topic_ID0EHROM.html
- OOXML custom XML data properties: https://ooxml.info/docs/22/22.5/

## Why this matters to Ascend

`package-action-proof/edge-fixture-policy` is the top blocker for the auditable mutation claim. The owner loop needs a repeatable scan, not a prose assertion, to decide whether the remaining generated signature and unknown-part proof cases are replaceable from the tracked public corpus.

## Probe/implementation

Added `fixtures/benchmarks/package-action-fixture-scan.ts` and tests. The scanner uses `git ls-files` by default so the proof is over tracked public fixtures, not ignored local fixture folders. It classifies fixtures by package feature:

- `docPropsCore`
- `docPropsCustom`
- `calcChain`
- `customXml`
- `macro`
- `chartOrDrawing`
- `signaturePackage`
- `syntheticUnknownPathFamily`

The release proof index now references this scan in the package-action `edge-fixture-policy` evidence.

## Results

`bun run fixtures/benchmarks/package-action-fixture-scan.ts --json` reports:

| Metric | Result |
| --- | ---: |
| Corpus | tracked git fixtures |
| Scanned XLSX/XLSM fixtures | 223 |
| Rejected fixtures | 1 |
| `docPropsCore` fixtures | 191 |
| `docPropsCustom` fixtures | 16 |
| `calcChain` fixtures | 52 |
| `customXml` fixtures | 4 |
| `macro` fixtures | 2 |
| `chartOrDrawing` fixtures | 46 |
| `signaturePackage` fixtures | 0 |
| `syntheticUnknownPathFamily` fixtures | 0 |

The result is `remaining-generated-edge-cases`. The docProps and calc-chain replacements are justified by tracked fixtures, but signature invalidation and unknown-part error remain generated structural cases unless product approves disclosed generated packages or acquires license-clear public binaries.

Validation:

```bash
bun test fixtures/benchmarks/package-action-fixture-scan.test.ts
bun run fixtures/benchmarks/package-action-fixture-scan.ts --json
```

## Confidence

High for the tracked-corpus result. Medium for the global public-fixture question: the scan does not prove that no license-clear public signed or unknown-part workbook exists elsewhere.

## Fold-in decision

Promote to product/release proof packaging as a reusable benchmark-only fixture scan. Keep `edge-fixture-policy` missing because the two remaining generated edge cases still lack tracked public replacements.

## Next question

Should product accept disclosed generated signature/unknown structural packages for guarded package-action proof, or should the next owner loop define a fixture acquisition policy for license-clear public binaries?
