# Package Action Current Proof Refresh

Date: 2026-05-15

## Question

Can the second-ranked "auditable package-part mutation" claim still be proven from existing plan/commit/package-action proof surfaces without adding a new mutation surface?

## Hypothesis

Yes. The tracked package-action proof harness should still cover `passthrough`, `regenerate`, `add`, `drop`, and `error`; align with rollback-journal package issues; and keep the honest boundary that this is local package evidence, not signed provenance or Excel semantic certification.

## External Sources Checked

- Open Packaging Conventions define package parts and relationships as the right proof unit for XLSX package accounting: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl documents preservation flags such as `keep_vba` while warning that unsupported workbook objects can be lost on save: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS CE write docs state that features outside documented support may not serialize: https://docs.sheetjs.com/docs/api/write-options/

## Why This Matters To Ascend

This is the release claim that makes safe mutation planning credible. The product-shaped claim is not "Ascend can write cells"; it is "Ascend can explain what happened to each relevant package part and surface review-required risk instead of silently claiming fidelity."

## Probe/Implementation

No production code changed. Reran existing proof and surface validations:

```bash
bun run fixtures/benchmarks/package-action-proof.ts
bun test fixtures/benchmarks/package-action-proof.test.ts
bun test packages/sdk/src/agent-workflow.test.ts -t "package action|package graph|journalSummary|compact commit"
bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow"
bun test apps/api/api.test.ts -t "plan and commit endpoints provide the safe write workflow"
bun test apps/mcp/src/index.test.ts -t "package action proof evidence"
```

Updated `research/experiments/syntheses/2026-05-package-action-proof-report.md` with the current proof table.

## Results

Latest proof run: 2026-05-15T03:56:57.904Z.

| Case | Commit actions | Digest pairs | Journal package issues | Proof issues | Proof ms | Post-write audits | Examples |
| --- | --- | ---: | ---: | ---: | ---: | --- | --- |
| docprops-passthrough | passthrough=4, regenerate=4, add=0, drop=0, error=0 | 8 | 1 | 0 | 0.197 | passed | passthrough:xl/workbook.xml; regenerate:xl/worksheets/sheet1.xml |
| regenerate-existing-sheet | passthrough=3, regenerate=5, add=0, drop=0, error=0 | 8 | 1 | 0 | 0.080 | passed | passthrough:xl/workbook.xml; regenerate:xl/styles.xml |
| add-sheet-part | passthrough=3, regenerate=5, add=1, drop=0, error=0 | 8 | 1 | 0 | 0.066 | passed | passthrough:xl/worksheets/sheet1.xml; regenerate:xl/workbook.xml; add:xl/worksheets/sheet2.xml |
| calc-chain-drop | passthrough=0, regenerate=5, add=0, drop=1, error=0 | 5 | 1 | 0 | 0.076 | passed | regenerate:xl/workbook.xml; drop:xl/calcChain.xml |
| signature-invalidation-drop | passthrough=1, regenerate=4, add=0, drop=2, error=0 | 5 | 1 | 0 | 0.089 | passed | passthrough:xl/workbook.xml; regenerate:xl/worksheets/sheet1.xml; drop:_xmlsignatures/origin.sigs |
| macro-passthrough | passthrough=6, regenerate=5, add=1, drop=0, error=0 | 11 | 1 | 0 | 0.228 | passed | passthrough:xl/workbook.xml; regenerate:xl/styles.xml; add:xl/sharedStrings.xml |
| chart-sidecar-accounting | passthrough=8, regenerate=6, add=1, drop=0, error=0 | 14 | 1 | 0 | 0.182 | passed | passthrough:xl/workbook.xml; regenerate:xl/styles.xml; add:xl/sharedStrings.xml |
| unknown-part-error | passthrough=2, regenerate=4, add=0, drop=0, error=1 | 7 | 1 | 1 | 0.073 | needs review | passthrough:xl/workbook.xml; regenerate:xl/worksheets/sheet1.xml; error:xl/custom/custom1.xml |

Combined commit actions: `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1`.

Surface validation passed for SDK plan/commit, CLI, API, and MCP package-action proof evidence.

## Confidence

High for guarded local package-action evidence. Medium for external publication because some edge cases remain code-generated synthetic packages and because the proof is not signed or tamper-evident.

## Fold-In Decision

Promote to product/correctness proof packaging only. Do not add another mutation surface. Keep the release index limited to safe-open and package-action proof artifacts until product defines publication storage and privacy policy.

## Next Question

Should the next loop refresh the release proof index digests for the two top proof artifacts, or stop and hand off the top two owner prompts now that both proofs are current?
