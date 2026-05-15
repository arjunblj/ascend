# Package Action Proof Surfaces

## Question

Can Ascend expose package action proof summaries through CLI/API/MCP plan and commit workflows so agents can see `passthrough`, `regenerate`, `add`, `drop`, and `error` evidence without calling SDK internals?

## Hypothesis

Yes. The SDK already builds `ascend-package-action-proof` objects from existing write-plan, write-policy, and package-graph audit evidence. Plan/commit surfaces can expose that proof behind an explicit opt-in flag without changing default payloads or writer behavior.

## External sources checked

- OpenTelemetry semantic conventions: https://opentelemetry.io/docs/concepts/semantic-conventions/
- OpenTelemetry logs data model: https://opentelemetry.io/docs/specs/otel/logs/data-model/
- Sigstore overview: https://docs.sigstore.dev/
- Sigstore signing overview: https://docs.sigstore.dev/cosign/signing/overview/
- In-toto Witness SLSA attestor docs: https://witness.dev/docs/docs/attestors/slsa/
- Microsoft Open XML formats and file extensions: https://support.microsoft.com/en-gb/office/open-xml-formats-and-file-name-extensions-5200d93c-3449-4380-8e11-31ef14555b18

## Why this matters to Ascend

Preservation-first XLSX edits need observable evidence about package parts. Existing release-proof work records package actions, but agents using CLI, API, or MCP could not request that evidence directly during normal safe plan/commit flows. External provenance and observability systems favor explicit typed evidence over ambiguous prose, which maps well to Ascend's package action taxonomy.

## Probe/implementation

- Inspected `createPackageActionProof()` and `createReleaseProofBundle()` in `packages/sdk/src/agent-workflow.ts`.
- Inspected CLI `plan`/`commit`, API `/plan`/`/commit`, and MCP `ascend.plan`/`ascend.commit`.
- Added CLI `--package-actions` to `ascend plan` and `ascend commit`.
- Added API request field `includePackageActions: true` for `/plan` and `/commit`.
- Added MCP parameter `includePackageActions` for `ascend.plan` and `ascend.commit`.
- The opt-in response field is `packageActions`, using the existing `ascend-package-action-proof` shape.
- Updated `docs/AGENT_API.md`, `docs/openapi.yaml`, `llms.txt`, `llms-full.txt`, and `agent-init` command guidance.

## Results

Focused validation passed:

```bash
bun test apps/cli/src/cli.test.ts -t "plan and commit implement safe agent workflow"
bun test apps/api/api.test.ts -t "plan and commit endpoints provide the safe write workflow"
bun test apps/mcp/src/index.test.ts -t "package action proof"
bunx biome check apps/cli/src/commands/plan.ts apps/cli/src/commands/commit.ts apps/cli/src/index.ts apps/cli/src/cli.test.ts apps/cli/src/commands/agent-init.ts apps/api/src/server.ts apps/api/api.test.ts apps/mcp/src/index.ts apps/mcp/src/index.test.ts docs/AGENT_API.md docs/openapi.yaml llms.txt llms-full.txt research/experiments/index.md research/experiments/runs/2026/2026-05-14-package-action-proof-surfaces/log.md
bunx tsc --build
bun run test:changed
```

The tests prove that CLI, API, and MCP plan/commit paths can return `packageActions.kind === "ascend-package-action-proof"` and nonzero regenerated package-part evidence.

`bun run test:changed` expanded to the full suite and passed with 4967 tests, 1 skip, and 0 failures.

## Confidence

High. This is an opt-in projection of an existing SDK proof object over already-tested plan/commit evidence. Defaults remain unchanged.

## Fold-in decision

Folded into correctness and product/DX loops. It makes package preservation evidence available to agents at the same workflow step where they inspect plan/commit safety.

## Next question

Can a columnar sidecar prototype be folded into a reusable benchmark harness that measures range/table scan acceleration without changing workbook semantics?
