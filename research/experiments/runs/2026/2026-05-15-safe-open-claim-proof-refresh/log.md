# Safe Open Claim Proof Refresh

## Question

Can the top-ranked "safe unknown workbook opening" claim be refreshed from existing SDK/CLI/API/MCP open-plan surfaces without adding a new product surface?

## Hypothesis

Yes. The current proof harness already measures package-feature routing before full hydration. The right work is to rerun it, verify public surfaces, tighten external contrast, and keep the remaining proof gaps explicit.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft Trusted Documents: https://support.microsoft.com/en-gb/office/trusted-documents-cf872bd8-47ec-4c02-baa5-1fdba1a11b53
- Microsoft Excel digital signatures and code signing: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing

## Why this matters to Ascend

The release claim board ranks safe unknown workbook opening as the top product/performance handoff. Ascend can credibly differentiate from generic workbook libraries by showing a cheap package-level decision point before workbook cells are hydrated, while being explicit that this is not malware scanning, sandboxing, or a trust decision.

## Probe/implementation

- Inspected the current safe-open report and proof harness:
  - `research/experiments/syntheses/2026-05-safe-open-proof-bundle.md`
  - `fixtures/benchmarks/safe-open-proof.ts`
  - `fixtures/benchmarks/safe-open-proof.test.ts`
- Left unrelated dirty benchmark-profiler edits unstaged.
- Reran the proof harness:
  - `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1`
  - `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1 --json`
- Reran surface validation:
  - `bun test fixtures/benchmarks/safe-open-proof.test.ts packages/sdk/src/open-plan.test.ts`
  - `bun test apps/cli/src/cli.test.ts -t "open-plan"`
  - `bun test apps/api/api.test.ts -t "open-plan"`
  - `bun test apps/mcp/src/index.test.ts -t "open_plan|open-plan"`
- Updated the safe-open proof bundle with fresh timings, validation status, and stronger OPC/Trusted Documents contrast.

## Results

The tracked proof harness still covers 9 cases: 6 public workbook fixtures, 2 synthetic package-risk fixtures, and 1 malformed byte case.

Latest markdown rerun:

| Case | Mode | Review | Risk families | Ratio |
| --- | --- | --- | --- | ---: |
| clean | formula | false | none | 8.67x |
| formula-heavy | formula | false | none | 34.45x |
| macro | metadata-only | true | preservedMacro | 22.82x |
| pivot | formula | false | none | 27.56x |
| ActiveX | metadata-only | true | preservedActiveX | 13.89x |
| chart | formula | false | none | 13.05x |
| signed | metadata-only | true | preservedSignature | 1.74x |
| unknown part | metadata-only | true | preservedOther | 2.27x |
| malformed | rejected | n/a | none | n/a |

SDK, CLI, API, and MCP validations passed. No new product surface was added.

## Confidence

High that the allowed wording is current: Ascend recommends a load mode and review branch from XLSX/XLSM package features before hydrating workbook cells. Medium for release headline readiness because signed and unknown package cases are still synthetic and the timings are local probe numbers, not thresholds.

## Fold-in decision

Promote to product/performance proof packaging only. Do not add another open-plan surface. The next owner should either replace synthetic signed/unknown cases with durable public fixture workbooks or explicitly publish them as generated proof cases.

## Next question

Can the package-action proof report receive the same no-new-surface refresh while preserving its honest boundaries around signatures, chart XML regeneration, and provenance?
