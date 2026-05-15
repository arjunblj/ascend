# Package Action Release Proof Packaging Rerun

## Question

Can the auditable package-part mutation owner produce a release-shaped proof report that preserves full SDK per-part evidence while using existing CLI/API/MCP package-action options only?

## Hypothesis

Yes. The existing package-action proof helpers and tracked benchmark harness already cover the action vocabulary. The next useful step is a current rerun, cross-surface validation, and report refresh with explicit boundaries.

## External sources checked

- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft `System.IO.Packaging` package parts, relationships, and signatures: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging
- Microsoft Excel digital signatures and invalidation on saved copies: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing
- Microsoft invalid-signature troubleshooting: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/document-contains-invalid-signatures-error
- openpyxl tutorial and unsupported-object preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options and writer boundaries: https://docs.sheetjs.com/docs/api/write-options/
- in-toto attestation model: https://github.com/in-toto/attestation

## Why this matters to Ascend

"Auditable package-part mutation" is the second top release handoff. Ascend's differentiator is not merely saving XLSX files; it is explaining which package parts were passed through, regenerated, added, dropped, or rejected, while keeping the boundary clear that this is local evidence rather than signed provenance or Excel semantic certification.

## Probe/implementation

- Ran `git status --short --branch`; an unrelated `packages/io-xlsx/src/reader/sheet.ts` parser optimization diff was present and left unstaged.
- Inspected the tracked proof harness in `fixtures/benchmarks/package-action-proof.ts`, the report, and package-action tests.
- Reran the proof harness:

```bash
bun run fixtures/benchmarks/package-action-proof.ts
```

- Validated the harness and existing SDK/CLI/API/MCP proof surfaces:

```bash
bun test fixtures/benchmarks/package-action-proof.test.ts
bun test packages/sdk/src/agent-workflow.test.ts -t "package action|package graph|journalSummary|compact commit"
bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow"
bun test apps/api/api.test.ts -t "plan and commit endpoints provide the safe write workflow"
bun test apps/mcp/src/index.test.ts -t "package action proof evidence"
```

- Updated `research/experiments/syntheses/2026-05-package-action-proof-report.md` with the fresh proof table and validation commands.

## Results

The proof rerun covered 8 cases:

- docProps passthrough;
- existing sheet regeneration;
- new sheet part addition;
- calc-chain drop;
- digital-signature invalidation/drop;
- macro project passthrough;
- chart/drawing sidecar accounting;
- unknown-part error.

Combined commit evidence covered every action kind:

- `passthrough=27`
- `regenerate=38`
- `add=3`
- `drop=3`
- `error=1`

Every case included source graph evidence, output digest evidence, expected action presence, and one rollback-journal package preservation issue. The unknown-part case intentionally ended with one package-action proof issue and post-write audit review, proving explicit `error` evidence rather than silent preservation.

## Confidence

High for guarded release language: Ascend can produce local package-part action evidence for representative mutation workflows. Medium for public release packaging because several edge cases are still synthetic packages; product may want durable binary public fixtures before publishing the claim externally.

## Fold-in decision

Promote to topic synthesis and correctness/product proof packaging. Do not add mutation surfaces. The next owner should decide release artifact storage/privacy and whether synthetic edge cases should become durable public fixture files.

## Next question

Can the token-bounded agent view claim be promoted to a product proof only after an end-to-end recovery example shows how an agent retrieves omitted evidence by locator under a strict token budget?
