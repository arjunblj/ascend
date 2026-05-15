# Safe Open Release Proof Rerun

## Question

Can the safe unknown workbook opening claim move from "handoff only" to a current release-proof artifact by rerunning the tracked harness, refreshing external contrast, and documenting validation without adding another product surface?

## Hypothesis

Yes. The existing `safe-open-proof` harness should still prove the claim from current code: package-feature routing happens before full workbook hydration, active/security/unknown package families route to metadata review, malformed bytes reject, and public-fixture open-plan timing remains materially cheaper than full hydration.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Safe Documents: https://support.microsoft.com/en-us/office/safe-documents-e2071599-fb31-442b-a30c-198c25e2aacd
- Microsoft Excel digital signatures and code signing: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing
- Microsoft invalid-signature behavior for signed workbooks: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/document-contains-invalid-signatures-error
- openpyxl tutorial and preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options and unsupported-feature boundary: https://docs.sheetjs.com/docs/api/write-options/

## Why this matters to Ascend

The claim board ranks safe unknown workbook opening as the top product/performance proof. A current rerun turns the tracked harness from stale research evidence into a claim board artifact with explicit validation and honest competitor contrast. It also prevents the loop from slipping into another implementation surface when the real missing work is proof packaging.

## Probe/implementation

- Inspected the tracked harness and synthesis:
  - `fixtures/benchmarks/safe-open-proof.ts`
  - `fixtures/benchmarks/safe-open-proof.test.ts`
  - `research/experiments/syntheses/2026-05-safe-open-proof-bundle.md`
- Reran the harness:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1
```

- Updated `research/experiments/syntheses/2026-05-safe-open-proof-bundle.md` with the current timings, Safe Documents boundary, and current validation status.
- Did not change SDK, CLI, API, MCP, or XLSX production code.

## Results

| Case | Fixture | Mode | Review before hydration | Risk families | Median open-plan ms | Median full-open ms | Full/open-plan ratio | Boundary |
| --- | --- | --- | --- | --- | ---: | ---: | ---: | --- |
| clean | `fixtures/xlsx/poi/SampleSS.xlsx` | formula | false | none | 0.185 | 1.930 | 10.42x | ok |
| formula-heavy | `fixtures/xlsx/poi/formula_stress_test.xlsx` | formula | false | none | 0.199 | 6.367 | 32.01x | ok |
| macro | `fixtures/xlsx/calamine/vba.xlsm` | metadata-only | true | preservedMacro | 0.094 | 1.481 | 15.74x | ok |
| pivot | `fixtures/xlsx/poi/ExcelPivotTableSample.xlsx` | formula | false | none | 0.161 | 2.280 | 14.16x | ok |
| ActiveX | `fixtures/xlsx/libreoffice/activex_checkbox.xlsx` | metadata-only | true | preservedActiveX | 0.092 | 1.740 | 18.87x | ok |
| chart | `fixtures/xlsx/poi/WithChart.xlsx` | formula | false | none | 0.108 | 1.417 | 13.14x | ok |
| signed | synthetic digital-signature package | metadata-only | true | preservedSignature | 0.050 | 0.086 | 1.73x | ok |
| unknown part | synthetic unknown package part | metadata-only | true | preservedOther | 0.037 | 0.081 | 2.20x | ok |
| malformed | synthetic malformed bytes | rejected | n/a | none | n/a | n/a | n/a | open-plan rejected: Missing end of central directory record |

Validation run:

```bash
bun test fixtures/benchmarks/safe-open-proof.test.ts
bun test packages/sdk/src/open-plan.test.ts
bun test apps/cli/src/cli.test.ts -t "open-plan"
bun test apps/api/api.test.ts -t "open-plan"
bun test apps/mcp/src/index.test.ts -t "open_plan"
bunx tsc --build
```

Markdown note: `bunx biome check` does not process these research markdown paths because they are ignored by the repository Biome configuration, so this checkpoint uses manual markdown review plus `git diff --check`.

## Confidence

High that the release-claim wording is current and proof-backed locally. Medium that timing numbers should be published as absolute values; they are release-environment evidence, not a CI threshold.

## Fold-in decision

Promote to product/performance proof packaging only. Do not add a new product surface. The next implementation loop should publish a compact proof report from the existing harness if product wants this as a release headline.

## Next question

Can auditable package-part mutation receive the same current release-proof rerun and report update without changing writer behavior?
