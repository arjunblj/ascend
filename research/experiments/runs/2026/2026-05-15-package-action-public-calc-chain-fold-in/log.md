# Package Action Public Calc-Chain Fold-In

## Question

Should `package-action-proof` replace its generated `calc-chain-drop` edge package with a checked-in public workbook fixture?

## Hypothesis

Yes. The prior edge-fixture scan showed many checked-in public workbooks with `xl/calcChain.xml`, and a focused probe proved `fixtures/xlsx/poi/Booleans.xlsx` can exercise the same drop action with post-write audits passing. This improves the second release claim without adding a new mutation surface.

## External sources checked

- Microsoft Open XML calculation-chain documentation says the calculation chain contains ordered formula-cell references for workbook calculation order: https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain
- Microsoft Excel calculation-chain metadata documentation describes Excel-created calculation-order metadata: https://support.microsoft.com/en-us/office/excel-calculation-chain-metadata-6e1b5819-6abd-4e94-bff5-838d4c576e01
- Open Packaging Conventions frame the proof as package-part accounting: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why this matters to Ascend

`package-action-proof/edge-fixture-policy` is still a product blocker. Moving calc-chain drop from generated topology to a public fixture reduces synthetic proof surface while keeping the honest boundary that signature and unknown-part cases remain generated.

## Probe/implementation

Changed `fixtures/benchmarks/package-action-proof.ts`:

- `calc-chain-drop` now uses `sourceKind: public-fixture`;
- fixture is `fixtures/xlsx/poi/Booleans.xlsx`;
- expected action remains `drop` for `xl/calcChain.xml`;
- no SDK, CLI, API, MCP, writer, or production behavior changed.

Updated package-action and release-proof-index tests for the new provenance mix.

## Results

Validation:

```bash
bun test fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/package-action-proof.ts fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bunx tsc --build
bun run test:changed
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

Results:

- focused tests: 8 pass, 0 fail;
- Biome passed on changed benchmark files;
- typecheck passed;
- `test:changed` passed: 4393 pass, 1 skip, 0 fail, 26755 expect calls;
- package-action source mix: 3 public fixtures, 2 generated workbooks, 3 generated edge packages;
- combined actions: `passthrough=32`, `regenerate=39`, `add=3`, `drop=3`, `error=1`;
- `calc-chain-drop` uses `fixtures/xlsx/poi/Booleans.xlsx`, drops `xl/calcChain.xml`, and passes post-write audits;
- release proof index still reports `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, and 9 missing owner requirements;
- package-action stable shape SHA-256 changed to `0f9eb22498bc528a63adc40e59a6acbbe07022fde6b2414fcbee73b8b3a56e41`.

## Confidence

High that calc-chain drop is now backed by a public fixture. High that the top release gate must remain blocked because signature invalidation and unknown-part error still rely on generated packages.

## Fold-in decision

Fold into the benchmark proof harness. This is a proof-quality improvement under the existing auditable package-part mutation claim, not a new product surface.

## Next question

Can `docprops-passthrough` be moved to a public fixture without losing the representative streaming proof, or should research focus next on release owner policy for the remaining signature/unknown generated cases?
