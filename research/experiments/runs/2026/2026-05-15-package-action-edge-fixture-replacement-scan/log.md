# Package Action Edge Fixture Replacement Scan

## Question

Can the package-action proof replace generated edge packages with checked-in public fixtures, or does the product gate still require explicit generated-fixture policy?

## Hypothesis

Partial replacement is possible. Public fixtures likely cover calc-chain and custom XML/document-property shapes, but not package signatures or arbitrary unknown package parts.

## External sources checked

- Microsoft calculation-chain documentation describes `calcChain` as SpreadsheetML calculation-order metadata: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain
- Microsoft Protected View keeps active/trust decisions separate from package accounting: https://support.microsoft.com/en-gb/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions frame packages as parts and relationships, matching per-part action proof: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl documents unsupported object preservation limits on save: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options document writer scope and data-preservation orientation: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

`package-action-proof/edge-fixture-policy` is a rank-10 owner blocker for the second release claim. Replacing generated edge cases with public fixtures would strengthen the claim, but silently swapping only easy cases could hide the harder signature and unknown-part boundaries.

## Probe/implementation

Ran the current compact package-action proof:

```bash
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
```

Scanned checked-in XLSX/XLSM package entries for edge families:

```bash
for f in $(find fixtures/xlsx -type f \( -name '*.xlsx' -o -name '*.xlsm' \)); do
  zipinfo -1 "$f" 2>/dev/null | rg '(^|/)(calcChain\.xml|customXml/|_xmlsignatures/|docProps/custom\.xml|vbaProject\.bin|charts?/|drawings?/)'
done
```

Then ran a focused public fixture probe with `fixtures/xlsx/poi/Booleans.xlsx`, editing a formula and generating commit package-action proof.

## Results

Current compact proof still has:

- 8 cases;
- 2 public fixtures;
- 2 generated workbooks;
- 4 generated edge packages;
- action totals `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1`;
- one representative streaming proof.

Checked-in fixture scan over 351 XLSX/XLSM files:

| Edge family | Public fixture count | Interpretation |
| --- | ---: | --- |
| `xl/calcChain.xml` | 101 | Public replacement likely available for `calc-chain-drop`. |
| `docProps/custom.xml` | 25 | Public custom document-property coverage exists, though not necessarily the same deterministic docProps passthrough shape. |
| `customXml/` | 8 | Public custom XML/package sidecar coverage exists. |
| `vbaProject.bin` | 2 | Macro proof already has public fixture coverage. |
| chart/drawing parts | 73 | Chart/drawing proof already has public fixture coverage. |
| `_xmlsignatures/` | 0 | No checked-in public signature package replacement. |
| `xl/custom/` or `custom/` unknown custom path | 0 | No checked-in equivalent for the synthetic unknown-part error case. |

Focused public calc-chain probe:

- input: `fixtures/xlsx/poi/Booleans.xlsx`;
- operation: set `Sheet1!B1` to formula `=A1+A1`;
- result: post-write audits passed;
- package action counts: `passthrough=5`, `regenerate=6`, `add=0`, `drop=1`, `error=0`;
- dropped part: `xl/calcChain.xml`.

## Confidence

High that calc-chain generated proof can be replaced by a public fixture in a future harness update. Medium that docProps/custom XML can be replaced without changing the precise proof semantics. High that signature and unknown-part cases remain generated or require acquisition policy.

## Fold-in decision

Promote to topic synthesis and future proof-harness prompt. Do not modify the package-action proof harness in this cycle because only one generated edge case has a proven public replacement and the owner gate still depends on signature/unknown policy.

## Next question

Should the next correctness/product loop swap `calc-chain-drop` to a public fixture while keeping signature and unknown cases generated and disclosed, or wait for a broader fixture acquisition policy?
