# SDK Release Smoke Tarball

## Question

Can Ascend prove a clean SDK external adoption path from a built local artifact without consumer overrides?

## Hypothesis

A narrow SDK-only release smoke can close part of the release packaging proof gap if it builds local library artifacts, bundles internal `@ascend/*` dependencies into an SDK tarball, installs only `@ascend/sdk` into a temp consumer app, and runs a plan/commit/reopen workflow.

## External sources checked

- npm `package.json` metadata and `files`/packaging behavior: https://docs.npmjs.com/cli/v10/configuring-npm/package-json
- npm local paths dependency syntax: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#local-paths
- npm `bundledDependencies`: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#bundleddependencies
- Bun package manager and install behavior: https://bun.sh/docs/cli/install
- Bun `pm pack`: https://bun.sh/docs/pm/cli/pack

## Why this matters to Ascend

The release proof bundle claim is not credible if external users need a repo checkout, workspace graph, or consumer-side overrides. This smoke proves one narrow SDK adoption path while keeping CLI/API/MCP packaging blockers open.

## Probe/implementation

Added `scripts/release-sdk-smoke.ts` and `release:sdk:smoke`.

The script:

1. Runs `bun run build:js`.
2. Copies built `packages/*/dist` artifacts into `/private/tmp/ascend-sdk-local-release/artifacts/packages`.
3. Rewrites internal artifact manifests for local package links.
4. Bundles internal library packages under the SDK artifact and writes `bundledDependencies`.
5. Packs `@ascend/sdk` into `/private/tmp/ascend-sdk-local-release/artifacts/ascend-sdk-0.0.0.tgz`.
6. Installs only that SDK tarball into a fresh temp consumer app.
7. Runs a smoke workflow: create workbook, save, open, plan `setCells`, commit, reopen, structural check, verify formula recalculation.

The smoke exposed a production post-write verification issue in packaged SDK usage: reopening output bytes as an `AscendWorkbook` did not preserve the document snapshot/source archive behavior needed by package-aware post-write summary. Fold-in:

- Added `WorkbookDocument.getWorkbookModel()` and `WorkbookDocument.writePlanSummary()`.
- Updated agent commit post-write verification to reopen output bytes through `WorkbookDocument.openPathSnapshot`.
- Preserved source archive bytes for planned-write summaries when available.

## Results

Validation command:

```bash
bun run release:sdk:smoke
```

Observed result:

- Built package JS artifacts.
- Packed `/private/tmp/ascend-sdk-local-release/artifacts/ascend-sdk-0.0.0.tgz`.
- Temp consumer installed only `@ascend/sdk` from the tarball.
- Smoke output included `planWouldSucceed: true`, `reopenedValid: true`, `B1=125`, and recalculated `C1=250`.

Focused validation also included:

```bash
bunx tsc --build
bunx biome check
bun run test:changed
```

## Confidence

Medium-high for SDK-only local artifact adoption. This does not prove npm publication, semver policy, docs packaging, CLI, API, MCP, provenance, or signed release artifacts.

## Fold-in decision

Promote to release/product loop as a reusable SDK packaging smoke. Keep the broader release proof bundle claim at `needs-one-more-fold-in` until CLI/API/MCP packaging, docs packaging, and release publication policy are handled.

## Next question

Can the release owner extend this smoke into packageable CLI/API/MCP artifacts without forcing the current TUI into headless installs?
