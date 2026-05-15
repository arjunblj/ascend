# 2026-05 Package Action Proof Report

Date: 2026-05-15

## Claim

Ascend can explain XLSX writes with local package-part action evidence: `passthrough`, `regenerate`, `add`, `drop`, and `error`.

## Claim Wording That Is Safe Today

Ascend can produce a local package-action proof for representative workbook mutations, showing which package parts were passed through, regenerated, added, dropped, or rejected with review-required errors. This is local evidence, not signed provenance, SLSA, in-toto, or Excel semantic certification.

## External Contrast

- [Microsoft OPC fundamentals](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview) define packages as parts plus relationships and describe package signatures as validation evidence over signed content.
- [Microsoft `System.IO.Packaging`](https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging) exposes package parts, relationships, and signatures as first-class concepts, matching Ascend's per-part proof boundary.
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/tutorial.html) documents `keep_vba` and warns that unsupported workbook objects may be lost on save.
- [SheetJS write options](https://docs.sheetjs.com/docs/api/write-options/) state that undocumented features may not serialize, which is a useful contrast for Ascend's explicit package-action accounting.
- [in-toto attestations](https://github.com/in-toto/attestation) are real supply-chain provenance artifacts; Ascend's package-action proof must not imply that level of signed attestation.

## Proof Bundle Status

| Required proof | Current evidence | Status |
| --- | --- | --- |
| Fixture mix | Synthetic docProps, new workbook edits, calc-chain, signature, unknown package part; public macro and chart workbooks | Covered for local proof; synthetic edge cases may still need durable public binaries before external publication |
| Action vocabulary | Combined commit evidence covers `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1` | Covered |
| Journal compatibility | Every proof case now reports one `package-part-preservation` journal issue alongside package-action evidence | Covered |
| SDK evidence shape | Tracked `fixtures/benchmarks/package-action-proof.ts` harness uses existing SDK plan/commit and package-action proof helpers | Covered |
| Validation gate | Harness test, prior full `test:changed`, typecheck, and Biome on changed TypeScript files | Covered in current loop |
| Competitor contrast | OPC, openpyxl, SheetJS, in-toto boundary | Covered |
| Honest boundary | Chart XML regenerates while drawing sidecars pass through; proof is local package evidence, not signed provenance or Excel recalc equivalence | Covered |

## Fresh Local Probe

Probe command:

```bash
bun run fixtures/benchmarks/package-action-proof.ts
```

Latest rerun: 2026-05-15T04:42:55.220Z.

| Case | Fixture | Input bytes | Output bytes | Commit actions | Source graph | Digest pairs | Journal package issues | Proof issues | Proof JSON bytes | Proof ms | Expected action present | Post-write audits | Examples |
| --- | --- | ---: | ---: | --- | --- | ---: | ---: | ---: | ---: | ---: | --- | --- | --- |
| docprops-passthrough | `synthetic docProps package` | 2286 | 3485 | passthrough=4, regenerate=4, add=0, drop=0, error=0 | true | 8 | 1 | 0 | 5155 | 0.152 | true | passed | passthrough:xl/workbook.xml; regenerate:xl/worksheets/sheet1.xml |
| regenerate-existing-sheet | `new Ascend workbook` | 4624 | 4707 | passthrough=3, regenerate=5, add=0, drop=0, error=0 | true | 8 | 1 | 0 | 5132 | 0.079 | true | passed | passthrough:xl/workbook.xml; regenerate:xl/styles.xml |
| add-sheet-part | `new Ascend workbook` | 4624 | 4512 | passthrough=3, regenerate=5, add=1, drop=0, error=0 | true | 8 | 1 | 0 | 5610 | 0.071 | true | passed | passthrough:xl/worksheets/sheet1.xml; regenerate:xl/workbook.xml; add:xl/worksheets/sheet2.xml |
| calc-chain-drop | `synthetic calcChain package` | 1776 | 2365 | passthrough=0, regenerate=5, add=0, drop=1, error=0 | true | 5 | 1 | 0 | 3760 | 0.079 | true | passed | regenerate:xl/workbook.xml; drop:xl/calcChain.xml |
| signature-invalidation-drop | `synthetic digital-signature package` | 2253 | 2058 | passthrough=1, regenerate=4, add=0, drop=2, error=0 | true | 5 | 1 | 0 | 4165 | 0.089 | true | passed | passthrough:xl/workbook.xml; regenerate:xl/worksheets/sheet1.xml; drop:_xmlsignatures/origin.sigs |
| macro-passthrough | `fixtures/xlsx/calamine/vba.xlsm` | 12752 | 12175 | passthrough=6, regenerate=5, add=1, drop=0, error=0 | true | 11 | 1 | 0 | 7359 | 0.204 | true | passed | passthrough:xl/workbook.xml; regenerate:xl/styles.xml; add:xl/sharedStrings.xml |
| chart-sidecar-accounting | `fixtures/xlsx/poi/WithChart.xlsx` | 10138 | 10899 | passthrough=8, regenerate=6, add=1, drop=0, error=0 | true | 14 | 1 | 0 | 9067 | 0.174 | true | passed | passthrough:xl/workbook.xml; regenerate:xl/styles.xml; add:xl/sharedStrings.xml |
| unknown-part-error | `synthetic unknown package part` | 1692 | 2315 | passthrough=2, regenerate=4, add=0, drop=0, error=1 | true | 7 | 1 | 1 | 4629 | 0.078 | true | needs review | passthrough:xl/workbook.xml; regenerate:xl/worksheets/sheet1.xml; error:xl/custom/custom1.xml |

Validation commands:

```bash
bun test fixtures/benchmarks/package-action-proof.test.ts
bun test packages/sdk/src/agent-workflow.test.ts -t "package action|package graph|journalSummary|compact commit"
bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow"
bun test apps/api/api.test.ts -t "plan and commit endpoints provide the safe write workflow"
bun test apps/mcp/src/index.test.ts -t "package action proof evidence"
```

Latest validation rerun passed on 2026-05-15:

- `bun test fixtures/benchmarks/package-action-proof.test.ts`
- `bun test packages/sdk/src/agent-workflow.test.ts -t "package action|package graph|journalSummary|compact commit"`
- `bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow"`
- `bun test apps/api/api.test.ts -t "plan and commit endpoints provide the safe write workflow"`
- `bun test apps/mcp/src/index.test.ts -t "package action proof evidence"`

## Interpretation

- The proof is strong enough for guarded release language around local package-part accounting.
- Package-action evidence and rollback-journal evidence now line up in the harness: every case has a package-preservation journal issue, and only the unknown-part case has a package-action proof issue.
- The proof should not say "chart byte passthrough." It should say chart/drawing package content is accounted for with per-part actions.
- The unknown-part case is intentionally not clean: it demonstrates explicit review-required evidence rather than silent preservation claims.
- Timing and JSON-size values are evidence for report shape, not performance thresholds.

## Fold-In Recommendation

Promote as the second release-proof artifact beside safe-open. Do not add a new mutation surface. The next product step is publication packaging: decide where the Markdown/JSON report lives and whether synthetic edge cases need public binary fixtures.

Release proof index status: `fixtures/benchmarks/release-proof-index.ts` now lists the exact reproduction command, publication blockers, owner-loop `readyWhen` gates, and one representative streaming writer proof for this artifact. It also carries a fail-closed `streaming-matrix-boundary` gate: do not describe the representative streaming proof as full streaming parity unless a performance owner approves that wording or the proof expands to streaming variants for every package-action scenario. The package-action artifact remains `needs-release-packaging` because synthetic edge packages must stay disclosed unless replaced by public binary fixtures, and this proof is local evidence rather than signed provenance or third-party attestation. The current no-timings stable shape digest is `9abebf576651551f58e00ccf8469d099b2c06dacd48391fe581a24e51a1e0afd`.

Do not track generated compact JSON yet. A local probe measured `7854` bytes for full generated JSON and `1896` bytes for a compact summary, but the current artifact shape includes generated metadata and duplicates harness assertions. Keep JSON as generated output until release-proof publication and privacy rules are explicit.
