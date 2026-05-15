# No-Op Preview Journal Surfaces

## Question

Should API and MCP preview/write accept explicit empty operation lists and return exact empty journals when agents request journal evidence?

## Hypothesis

Yes. An explicit `ops: []` is an intentional no-op batch, not malformed input. If `journal: true` is set, the result should carry the same exact empty journal shape as SDK no-op apply/preview, so agents can audit "nothing changed" without special-casing transport errors.

## External sources checked

- JSON Patch RFC 6902: https://www.rfc-editor.org/rfc/rfc6902
- HTTP PATCH RFC 5789: https://www.rfc-editor.org/rfc/rfc5789
- OpenAPI request body docs: https://spec.openapis.org/oas/latest.html

## Why this matters to Ascend

Auditable package-part mutation is not only about changed workbooks. Agents also need trustworthy no-op evidence when a generated plan compiles to zero operations or when a guarded workflow decides no mutation is needed. Returning a transport error for `ops: []` makes empty journals unreliable across API/MCP surfaces even though SDK operation parsing already accepts explicit empty ops.

## Probe/implementation

Finished the in-flight API/MCP tests and fixed the SDK preview path:

- `AscendWorkbook.preview([],{ journal: true })` now returns a successful empty preview with the exact empty mutation journal instead of calling the operation engine with no ops.
- API `/preview` and `/write` tests assert explicit `ops: []` plus `journal: true` returns the full exact empty journal shape.
- MCP `ascend.preview` and `ascend.write` tests assert the same shape through structured tool responses.

## Results

Focused validation passed:

```bash
bun test apps/api/src/server.test.ts -t "preview and write return exact empty journals for no-op requests"
bun test apps/mcp/src/index.test.ts -t "ascend.preview and ascend.write return exact empty journals for no-op requests"
bun test packages/sdk/src/interactive-contract.test.ts -t "journal"
```

The SDK journal-focused slice passed 140 tests with 1081 assertions.

## Confidence

High for the no-op journal behavior across SDK/API/MCP. Medium for broader agent workflows until the full changed-test gate runs after this patch.

## Fold-in decision

Promote to correctness loop as a tiny scoped production fix. This supports the auditable mutation claim but does not add a new public surface or new claim wording.

## Next question

Can safe-open latency evidence be rerun from a tracked-clean worktree so the performance owner gets useful timing data without turning a dirty local probe into a release claim?
