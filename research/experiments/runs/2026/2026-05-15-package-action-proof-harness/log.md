# Package Action Proof Harness

## Question

Can Ascend generate a tracked, fixture-backed proof report for auditable package-part mutation across `passthrough`, `regenerate`, `add`, `drop`, and `error` without expanding default product payloads?

## Hypothesis

Yes. The existing SDK plan/commit workflow and package action proof helper already expose enough evidence. A benchmark/proof harness can run representative public and code-generated workbook workflows, assert expected part actions, measure proof JSON size and helper overhead, and keep the output as release-claim evidence rather than a new API/CLI/MCP surface.

## External sources checked

- SheetJS CE write options document writer scope, data-preservation orientation, and explicit `bookVBA` preservation knobs: https://docs.sheetjs.com/docs/api/write-options/
- openpyxl usage docs document `keep_vba` and warn that unsupported workbook objects such as images/charts can be lost on save: https://openpyxl.readthedocs.io/en/3.0/usage.html
- Microsoft `System.IO.Packaging` documents packages as parts plus relationships and digital signatures over parts/relationships: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging
- Microsoft `Package` docs expose create/delete APIs for parts and relationships, matching Ascend's package-part action boundary: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging.package

## Why this matters to Ascend

The release claim board says "auditable package-part mutation" still needs fixture-backed proof across the full action vocabulary. After the source-graph fold-in, commit proofs can distinguish `add` from `regenerate`, but the evidence was still spread across focused tests and logs. A tracked proof harness gives product/correctness loops one repeatable artifact for the claim while preserving the boundary that this is local package evidence, not signed provenance or Excel semantic equivalence.

## Probe/implementation

- Inspected `createPackageActionProof()`, `createAgentCommitPackageActionProof()`, write-policy/package-graph tests, and public fixture inventory.
- Added `fixtures/benchmarks/package-action-proof.ts`.
  - Runs existing SDK `createAgentPlan()` and `commitAgentPlan()` workflows.
  - Emits Markdown by default and JSON with `--json`.
  - Asserts expected commit actions per case before reporting success.
  - Measures package-action proof JSON size and helper overhead.
- Added `fixtures/benchmarks/package-action-proof.test.ts`.
  - Asserts the release-claim scenario list.
  - Asserts combined commit evidence covers all five actions without timing thresholds.
  - Asserts report wording keeps honest boundaries.

Cases:

- synthetic docProps package: docProps passthrough
- new Ascend workbook value edit: existing sheet regeneration
- new Ascend workbook add sheet: added worksheet part
- synthetic calcChain package: calc-chain drop
- synthetic signed package: signature invalidation drop
- public `fixtures/xlsx/calamine/vba.xlsm`: macro passthrough
- public `fixtures/xlsx/poi/WithChart.xlsx`: chart/drawing sidecar accounting
- synthetic unknown package part: post-write audit error

## Results

Local proof command:

```bash
bun run fixtures/benchmarks/package-action-proof.ts
```

| Case | Fixture | Input bytes | Output bytes | Commit actions | Source graph | Digest pairs | Issues | Proof JSON bytes | Proof ms | Expected action present | Post-write audits |
| --- | --- | ---: | ---: | --- | --- | ---: | ---: | ---: | ---: | --- | --- |
| docprops-passthrough | synthetic docProps package | 2286 | 3698 | passthrough=4, regenerate=4, add=0, drop=0, error=0 | true | 8 | 0 | 5155 | 0.145 | true | passed |
| regenerate-existing-sheet | new Ascend workbook | 4624 | 4707 | passthrough=3, regenerate=5, add=0, drop=0, error=0 | true | 8 | 0 | 5132 | 0.073 | true | passed |
| add-sheet-part | new Ascend workbook | 4624 | 4512 | passthrough=3, regenerate=5, add=1, drop=0, error=0 | true | 8 | 0 | 5610 | 0.060 | true | passed |
| calc-chain-drop | synthetic calcChain package | 1776 | 2365 | passthrough=0, regenerate=5, add=0, drop=1, error=0 | true | 5 | 0 | 3760 | 0.069 | true | passed |
| signature-invalidation-drop | synthetic digital-signature package | 2253 | 2271 | passthrough=1, regenerate=4, add=0, drop=2, error=0 | true | 5 | 0 | 4165 | 0.076 | true | passed |
| macro-passthrough | `fixtures/xlsx/calamine/vba.xlsm` | 12752 | 13751 | passthrough=6, regenerate=5, add=1, drop=0, error=0 | true | 11 | 0 | 7359 | 0.181 | true | passed |
| chart-sidecar-accounting | `fixtures/xlsx/poi/WithChart.xlsx` | 10138 | 12477 | passthrough=8, regenerate=6, add=1, drop=0, error=0 | true | 14 | 0 | 9067 | 0.156 | true | passed |
| unknown-part-error | synthetic unknown package part | 1692 | 2614 | passthrough=2, regenerate=4, add=0, drop=0, error=1 | true | 7 | 1 | 4629 | 0.064 | true | needs review |

Combined commit actions: passthrough=27, regenerate=38, add=3, drop=3, error=1.

Dead-end/boundary found:

- The chart fixture does not support a simple "chart sidecar passthrough" claim. Ascend regenerates `xl/charts/chart1.xml` and passes drawing sidecars through. The claim should say chart/drawing package content is accounted for with per-part actions, not that every chart part is passed through byte-for-byte.

Validation passed:

- `bun test fixtures/benchmarks/package-action-proof.test.ts`
- `bunx biome check fixtures/benchmarks/package-action-proof.ts fixtures/benchmarks/package-action-proof.test.ts`
- `bunx tsc --build`
- `bun run test:changed` (5000 pass, 1 skip)

## Confidence

High that the harness proves the action taxonomy and catches overbroad claim wording. Medium that the fixture set is enough for a public release claim: signed, calc-chain, docProps, and unknown cases are durable generated packages, but product may still want binary public fixtures before publishing external proof numbers.

## Fold-in decision

Fold into the correctness/product proof loop as a tracked benchmark/proof harness. Do not add another SDK/CLI/API/MCP surface. Update claim wording from "drawing/chart sidecar passthrough" to "chart/drawing sidecar accounting" unless a future writer change proves chart XML byte passthrough.

## Next question

Can the release proof bundle consume the safe-open and package-action proof harness outputs as local evidence artifacts without claiming signed provenance?
