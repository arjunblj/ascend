# 2026-05 Safe Open Proof Bundle

Date: 2026-05-15

## Claim

Ascend can route unknown XLSX/XLSM files through a package-level open plan before full workbook hydration, recommending a load mode and review step from observable package features.

## Claim Wording That Is Safe Today

Ascend recommends a load mode and trust-review branch from XLSX package features before hydrating workbook cells. This is pre-hydration risk routing, not malware scanning, sandboxing, or a guarantee that workbook content is safe.

## External Contrast

- [Microsoft Protected View](https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653) opens potentially unsafe files read-only or in Protected View. Ascend's OSS claim is narrower: package-feature routing before choosing hydration mode.
- [Microsoft Safe Documents](https://support.microsoft.com/en-us/office/safe-documents-e2071599-fb31-442b-a30c-198c25e2aacd) uses Microsoft Defender scanning for documents opened in Protected View. Ascend must not imply equivalent threat detection.
- [Microsoft Excel digital signatures](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing) help prove a workbook has not changed since signing, and saving after modification invalidates that signature. Ascend should route signature material to review before edit planning.
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/tutorial.html) documents `keep_vba`, but warns that not all Excel items are read and unsupported shapes can be lost on save. Ascend should contrast with explicit preservation/risk routing, not broad compatibility.
- [SheetJS write options](https://docs.sheetjs.com/docs/api/write-options/) note that features outside documented support may not serialize. Ascend's proof should show when package features force metadata review rather than silently assuming a full semantic model.

## Proof Bundle Status

| Required proof | Current evidence | Status |
| --- | --- | --- |
| Fixture mix | Public clean, formula-heavy, macro, pivot, ActiveX, and chart fixtures; synthetic digital signature, unknown package part, and malformed bytes | Covered for local proof; signed/unknown/malformed should become fixture-backed if promoted |
| Benchmark | Tracked `fixtures/benchmarks/safe-open-proof.ts` harness measures open-plan against full hydration and renders Markdown/JSON proof output | Strong enough for product proof direction, not a CI threshold |
| API/CLI/MCP surface | Existing SDK `inspectWorkbookOpenPlan`, CLI `ascend open-plan`, API `POST /open-plan`, MCP `ascend.open_plan` | Implemented; do not add another surface |
| Validation gate | Focused open-plan tests, API/CLI/MCP open-plan tests, tracked proof harness tests, typecheck, and markdown diff check | Covered in the current rerun; no production behavior changed |
| Competitor contrast | Microsoft Protected View, Microsoft Safe Documents, Excel digital signatures, openpyxl, SheetJS | Covered |
| Honest boundary | Malformed bytes reject; active/security/unknown features route to metadata-only review; no malware/sandbox claim | Covered |

## Fresh Local Probe

Probe command:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1
```

The proof harness is tracked and intentionally not a new product surface. It generates durable synthetic signed, unknown-part, and malformed cases in code, and uses public workbook fixtures for real-workbook cases.

| Case | Fixture | Bytes | Mode | Review before hydration | Risk families | Parts | Relationships | Median open-plan ms | Median full-open ms | Full/open-plan ratio | Boundary |
| --- | --- | ---: | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| clean | `fixtures/xlsx/poi/SampleSS.xlsx` | 9112 | formula | false | none | 13 | 10 | 0.185 | 1.930 | 10.42x | ok |
| formula-heavy | `fixtures/xlsx/poi/formula_stress_test.xlsx` | 64769 | formula | false | none | 27 | 22 | 0.199 | 6.367 | 32.01x | ok |
| macro | `fixtures/xlsx/calamine/vba.xlsm` | 12752 | metadata-only | true | preservedMacro | 12 | 9 | 0.094 | 1.481 | 15.74x | ok |
| pivot | `fixtures/xlsx/poi/ExcelPivotTableSample.xlsx` | 19460 | formula | false | none | 27 | 19 | 0.161 | 2.280 | 14.16x | ok |
| ActiveX | `fixtures/xlsx/libreoffice/activex_checkbox.xlsx` | 12433 | metadata-only | true | preservedActiveX | 17 | 12 | 0.092 | 1.740 | 18.87x | ok |
| chart | `fixtures/xlsx/poi/WithChart.xlsx` | 10138 | formula | false | none | 15 | 10 | 0.108 | 1.417 | 13.14x | ok |
| signed | synthetic digital-signature package | 2254 | metadata-only | true | preservedSignature | 8 | 4 | 0.050 | 0.086 | 1.73x | ok |
| unknown part | synthetic unknown package part | 1697 | metadata-only | true | preservedOther | 6 | 3 | 0.037 | 0.081 | 2.20x | ok |
| malformed | synthetic malformed bytes | 9 | rejected | n/a | n/a | n/a | n/a | n/a | n/a | n/a | open-plan rejected: Missing end of central directory record |

## Interpretation

- Public workbook cases show open-plan as a cheap pre-hydration routing step: 10.42x to 32.01x faster than full hydration in this local probe.
- Active content and security-sensitive package material route to `metadata-only` with `reviewBeforeHydration: true`.
- Unknown package material routes to `metadata-only` review instead of pretending the workbook is fully understood.
- Malformed bytes do not get a recommendation. They reject at package inspection, which should be presented as a boundary in any proof bundle.
- Synthetic signed and unknown cases are now durable code-generated package cases in the tracked harness, but the product loop may still choose to add binary public fixtures before publishing.

## Product Boundary

Do claim:

- pre-hydration package-feature routing;
- recommended load mode for caller intent;
- explicit `reviewBeforeHydration` for active-content, security, and unknown package families;
- lower routing cost than full hydration on public fixtures.

Do not claim:

- malware detection;
- Microsoft Safe Documents or Defender-equivalent scanning;
- sandboxed opening;
- trusted active-content execution;
- complete Excel compatibility;
- that malformed packages can always be inspected;
- that local latency numbers are a benchmark threshold.

## Fold-In Recommendation

Promote to product/performance as a proof bundle over existing surfaces. Do not add another CLI/API/MCP surface. The tracked harness is the repeatable report generator; the next production-sized step should only package the report output into release materials if product wants this claim published.

## Next Handoff

```text
/goal Publish the safe unknown workbook opening proof bundle without adding a new product surface. Turn the current local proof into a repeatable tracked report over public fixtures, replacing synthetic signed/unknown cases with durable fixture workbooks if possible. Show package fingerprint, recommended load mode, reviewBeforeHydration, risk families, package counts, malformed rejection, and latency versus full hydration. Keep boundaries explicit: this is pre-hydration package routing, not malware scanning or sandboxing.
```
