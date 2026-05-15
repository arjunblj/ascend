# Safe Open Release Proof Packaging Rerun

## Question

Can the safe unknown workbook opening owner produce a release-proof report from existing SDK/CLI/API/MCP open-plan surfaces and public fixtures, without adding a new command, endpoint, or MCP tool?

## Hypothesis

Yes. The existing `inspectWorkbookOpenPlan` implementation and tracked `fixtures/benchmarks/safe-open-proof.ts` harness already provide the proof shape. The useful work is a fresh rerun, cross-surface validation, and tighter release-boundary language.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Office file validation and Protected View backport context: https://www.microsoft.com/en-us/msrc/blog/2010/12/more-about-the-office-file-validation-backport-plan
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft ActiveX settings in Office files: https://support.microsoft.com/en-us/office/enable-or-disable-activex-settings-in-office-files-f1303e08-a3f8-41c5-a17e-b0b8898743ed
- Microsoft macro security settings in Excel: https://support.microsoft.com/en-gb/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373
- openpyxl tutorial and preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options and writer boundaries: https://docs.sheetjs.com/docs/api/write-options/
- SheetJS VBA feature notes: https://docs.sheetjs.com/docs/csf/features/vba/

## Why this matters to Ascend

"Safe unknown workbook opening" is currently the top product/performance handoff in the release claim board. To make the claim credible, Ascend needs proof over observable package features before hydration: fixture coverage, timing versus full hydration, public surface evidence, validation commands, competitor contrast, and explicit safety boundaries.

## Probe/implementation

- Ran `git status --short --branch`; left unrelated `packages/sdk/src/sheet-handle.ts` compact-read keying diff unstaged.
- Inspected the current open-plan implementation in `packages/sdk/src/open-plan.ts`, the CLI/API/MCP open-plan surfaces, and the tracked safe-open proof harness.
- Reran the proof harness:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1 --json
```

- Validated the harness and public surfaces:

```bash
bun test fixtures/benchmarks/safe-open-proof.test.ts packages/sdk/src/open-plan.test.ts
bun test apps/cli/src/cli.test.ts -t "open-plan"
bun test apps/api/api.test.ts -t "open-plan"
bun test apps/mcp/src/index.test.ts -t "open_plan|open-plan"
```

- Updated `research/experiments/syntheses/2026-05-safe-open-proof-bundle.md` with the fresh rerun table.

## Results

The proof rerun covered 9 cases:

- 6 public workbook fixtures: clean, formula-heavy, macro, pivot, ActiveX, chart.
- 2 synthetic package-edge fixtures: signed package and unknown part.
- 1 malformed-byte boundary case.

Routing results:

- clean, formula-heavy, pivot, and chart routed to `formula` without review.
- macro routed to `metadata-only` with `reviewBeforeHydration: true` and `preservedMacro`.
- ActiveX routed to `metadata-only` with `reviewBeforeHydration: true` and `preservedActiveX`.
- synthetic signed routed to `metadata-only` with `reviewBeforeHydration: true` and `preservedSignature`.
- synthetic unknown part routed to `metadata-only` with `reviewBeforeHydration: true` and `preservedOther`.
- malformed bytes rejected with `open-plan rejected: Missing end of central directory record`.

Timing results from the local rerun:

- Public workbook full-open/open-plan ratios ranged from `9.14x` to `31.85x`.
- Synthetic edge cases remained much smaller, with ratios from `1.84x` to `2.26x`.
- These are proof-run observations, not release thresholds.

Surface validation passed for SDK, CLI, API, and MCP open-plan behavior. No production surface was added.

## Confidence

High that Ascend can claim pre-hydration package-feature routing in guarded language. Medium for release-publication readiness because signed and unknown-part proof still relies on synthetic code-generated packages; product may want durable public binary fixtures before publishing the claim externally.

## Fold-in decision

Promote to topic synthesis and product/performance proof packaging. Do not add production surfaces. The next owner should package the existing report into release materials and decide whether to replace synthetic signed/unknown cases with public binary fixtures.

## Next question

Can the auditable package-part mutation owner produce an equally release-shaped proof report that preserves full SDK per-part evidence while exposing compact CLI/API/MCP summaries through existing flags/options only?
