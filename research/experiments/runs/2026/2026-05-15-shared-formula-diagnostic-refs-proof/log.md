# Shared Formula Diagnostic Refs Proof

## Question

When an imported shared-formula member is corrupted, do compact agent plan checks and write-policy diagnostics preserve the exact affected workbook refs?

## Hypothesis

Yes. The real shared-formula trust fixture should prove both the compact check issue and the blocking write-policy diagnostic carry `Label!A3` and `Label!A2`, so agents can explain the failure without reopening the full workbook.

## External sources checked

- Language Server Protocol diagnostics use structured fields and ranges so tools can locate issues instead of relying on prose: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- OpenTelemetry semantic conventions reinforce stable structured event attributes for downstream tooling: https://opentelemetry.io/docs/specs/otel/semantic-conventions/
- HyperFormula documents formula references and dependency graph behavior, useful competitor context for formula trust diagnostics: https://hyperformula.handsontable.com/guide/cell-references.html

## Why this matters to Ascend

The auditable mutation claim is stronger when an agent can see not just that a write is blocked, but the exact imported formula-binding refs involved. This is proof support for trustworthy mutation planning, not a new formula surface.

## Probe/implementation

Finished the in-flight SDK test assertion in `packages/sdk/src/agent-workflow.test.ts`:

- captures `compactAgentPlanResult(plan)` for the real imported shared-formula corruption case;
- asserts the compact check issue includes refs `Label!A3` and `Label!A2`;
- asserts `compact.writePolicy.diagnostics` carries the same refs inside `pre-write-check-error` details.

Commands run:

```bash
bun test packages/sdk/src/agent-workflow.test.ts
bunx biome check packages/sdk/src/agent-workflow.test.ts
bunx tsc --build
```

## Results

- SDK agent workflow tests passed: 74 tests, 497 assertions.
- Biome passed for the touched SDK test file.
- `bunx tsc --build` passed.

## Confidence

High for the compact diagnostic ref proof. The assertion runs on a real imported shared-formula fixture rather than only on synthetic in-memory metadata.

## Fold-in decision

Promote to correctness loop as test-only diagnostic hardening. Do not promote formula rename or a new diagnostic API.

## Next question

Should release-proof owner handoff treat exact diagnostic refs as a proof requirement for auditable package-part mutation, or keep it as supporting correctness hygiene?
