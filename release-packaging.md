# Release Packaging Audit

Date: 2026-05-15

## SDK Packaging Implementation Update

The first SDK release-packaging target now passes.

Checked-in command:

```bash
bun run release:sdk:smoke
```

What it does:

1. Builds the SDK library JS artifacts.
2. Produces a local release tarball at `/private/tmp/ascend-sdk-local-release/artifacts/ascend-sdk-0.0.0.tgz`.
3. Creates a fresh temp consumer app at `/private/tmp/ascend-sdk-local-release/consumer`.
4. Installs only `@ascend/sdk` from the produced tarball, with no consumer overrides and no direct dependencies on internal `@ascend/*` packages.
5. Runs the SDK workflow: create/open workbook, inspect, plan, commit, reopen, and check.

Validation result:

```text
SDK release smoke passed: /private/tmp/ascend-sdk-local-release/consumer
SDK artifact: /private/tmp/ascend-sdk-local-release/artifacts/ascend-sdk-0.0.0.tgz
```

The SDK tarball currently bundles Ascend internal packages inside the local SDK artifact so the temp app does not need repo-local package graph knowledge. Public runtime dependencies remain normal package dependencies.

## Headless CLI Split Update

The CLI no longer hard-depends on the current TUI package at startup.

Changed shape:

- `apps/cli/package.json` no longer declares `@ascend/tui`.
- `apps/cli/tsconfig.json` no longer references `apps/tui`.
- `ascend tui` and `ascend open` now lazy-load `@ascend/tui` only when those commands are invoked.
- Headless commands such as `inspect`, `plan`, `commit`, and `check` can load the CLI without resolving TUI dependencies.

Validation:

```bash
bun test apps/cli/src/cli.test.ts
bun run build
```

## App Dist And Bin Update

CLI, API, and MCP now get real bundled app JS and generated publish manifests from `bun run build`.

Generated app artifacts:

| Package | Dist entry | Generated bin | Import/startup validation |
| --- | --- | --- | --- |
| `@ascend/cli` | `apps/cli/dist/index.js` | `ascend` | `bun apps/cli/dist/index.js --version` |
| `@ascend/api` | `apps/api/dist/index.js` | `ascend-api` | import `createApiFetch/createServer`; `/health` returns `ok` |
| `@ascend/mcp` | `apps/mcp/dist/index.js` | `ascend-mcp` | import `createServer`; `32` MCP tools registered |

The API, MCP, and CLI entrypoints use an explicit direct-run check instead of relying on `import.meta.main`, so importing built app packages does not accidentally start a process.

## App External Install Smoke Update

The CLI/API/MCP app release-packaging target now passes from produced local app tarballs.

Checked-in command:

```bash
bun run release:apps:smoke
```

What it does:

1. Builds the release JS artifacts.
2. Produces local release tarballs at `/private/tmp/ascend-apps-local-release/artifacts/ascend-{cli,api,mcp}-0.0.0.tgz`.
3. Creates a fresh temp consumer app at `/private/tmp/ascend-apps-local-release/consumer`.
4. Installs only `@ascend/cli`, `@ascend/api`, and `@ascend/mcp` from the produced tarballs, with no consumer overrides and no direct dependencies on internal `@ascend/*` packages.
5. Runs the workbook workflow through the installed CLI bin, installed API exports, and installed MCP tool handlers: create/setup, inspect, plan, commit, reopen/check, and read/verify `B1 = 125` and `C1 = 250`.
6. Verifies installed docs through `ascend docs "plan commit" --json` and `ascend.search_docs`.

Validation result:

```text
App release smoke passed: /private/tmp/ascend-apps-local-release/consumer
cli artifact: /private/tmp/ascend-apps-local-release/artifacts/ascend-cli-0.0.0.tgz
api artifact: /private/tmp/ascend-apps-local-release/artifacts/ascend-api-0.0.0.tgz
mcp artifact: /private/tmp/ascend-apps-local-release/artifacts/ascend-mcp-0.0.0.tgz
```

## Installed Docs Update

SDK agent docs/examples now work from the installed SDK package.

Changed shape:

- `packages/sdk/src/agent-docs.ts` resolves package-local docs first, then falls back to the repo root for source-tree development.
- `scripts/build-packages.ts` copies the known `llms`, `docs`, and `examples` assets into `packages/sdk/dist`.
- `scripts/release-sdk-smoke.ts` verifies installed-package docs by calling `readAgentDoc('llms.txt')` and `searchAgentDocs({ query: 'plan commit' })` from the fresh temp consumer app.

Validation:

```bash
bun test packages/sdk/src/agent-docs.test.ts
bun run release:sdk:smoke
```

## Verdict

Ascend is no longer blocked on the first SDK external install smoke, installed SDK docs, the current-TUI startup split, real CLI/API/MCP dist JS and bins, or packaged CLI/API/MCP external install smokes.

The SDK workflow now works from a temp external app with a single local `@ascend/sdk` tarball and no consumer overrides. The installed SDK package can also read/search bundled agent docs. CLI/API/MCP now build real app entrypoints and publish manifests, and a fresh external temp app can install their produced local tarballs and run the same workbook workflow without repo-local context.

Previous audit stop condition was reached; the current implementation sequence has closed the SDK-first path, headless CLI split, app dist/bin artifacts, installed SDK docs, and external app package smokes. Final registry/publish topology remains intentionally deferred.

## Manifest, Export, Bin, And Build Audit

Source manifests:

| Package | Current source manifest state | Release risk |
| --- | --- | --- |
| `@ascend/schema` | `private: true`, `main: src/index.ts`, no `exports`, internal packages use `workspace:*` | Cannot publish source manifest. |
| `@ascend/core` | `private: true`, `main: src/index.ts`, no `exports`, `@ascend/schema: workspace:*` | Cannot publish source manifest. |
| `@ascend/formulas` | `private: true`, `main: src/index.ts`, no `exports`, workspace deps | Cannot publish source manifest. |
| `@ascend/engine` | `private: true`, `main: src/index.ts`, no `exports`, workspace deps | Cannot publish source manifest. |
| `@ascend/io-xlsx` | `private: true`, `main: src/index.ts`, no `exports`, workspace deps | Cannot publish source manifest. |
| `@ascend/io-csv` | `private: true`, `main: src/index.ts`, no `exports`, workspace deps | Cannot publish source manifest. |
| `@ascend/verify` | `private: true`, `main: src/index.ts`, no `exports`, workspace deps | Cannot publish source manifest. |
| `@ascend/sdk` | `private: true`, `main: src/index.ts`, no `exports`, workspace deps | Cannot publish source manifest. |
| `@ascend/cli` | `private: true`, `main: src/index.ts`, `bin.ascend: src/index.ts`, depends on `@ascend/tui: workspace:*` | Headless CLI package is blocked by source bin and hard TUI dependency. |
| `@ascend/api` | `private: true`, `main: src/index.ts`, no `exports`, no `bin`, workspace deps | No packageable API server entrypoint. |
| `@ascend/mcp` | `private: true`, `main: src/index.ts`, no `exports`, no `bin`, workspace deps | No packageable MCP server entrypoint. |

Library build path:

- `bun run build` passed.
- `scripts/build-packages.ts` only builds `schema`, `core`, `formulas`, `engine`, `io-xlsx`, `io-csv`, `verify`, and `sdk`.
- Library `dist/package.json` files are publish-style: `private: false`, `main: ./index.js`, root `exports`, `types: ./index.d.ts`, and internal deps rewritten from `workspace:*` to `0.0.0`.
- App `dist` folders have declarations, but no `dist/package.json`. Their current `dist/index.js` files are placeholders:
  - `apps/cli/dist/index.js`: `console.log("ascend cli")`
  - `apps/api/dist/index.js`: `console.log("ascend api")`
  - `apps/mcp/dist/index.js`: `console.log("ascend mcp")`
- App `tsconfig.json` files are `emitDeclarationOnly: true`, so the releasable JS for CLI/API/MCP is not built.

Implementation update: `scripts/build-packages.ts` now builds CLI/API/MCP runtime JS and writes app `dist/package.json` files with bins. The audit bullets above describe the original blocker state.

Missing export/bin gaps:

- Library `dist` manifests export only `"."`. That is acceptable for SDK-first use, but not for documented deep surfaces unless we intentionally add subpath exports.
- Implementation update: API `dist` exposes `createApiFetch` and `createServer` through the generated root export.
- Implementation update: MCP `dist` exposes `createServer` through the generated root export and `ascend-mcp` bin.
- Implementation update: CLI `dist` exposes the `ascend` bin, and headless startup no longer resolves `@ascend/tui`.

Docs/OpenAPI:

- `bun test packages/sdk/src/agent-docs.test.ts` passed: OpenAPI currently lists implemented endpoints and docs/examples vocabulary is aligned inside the repo.
- Implementation update: external SDK package docs are now packaged and validated by `bun run release:sdk:smoke`; the installed smoke returned `docHits: 5`.

## Temp External Install Path

Implemented SDK path: `/private/tmp/ascend-sdk-local-release`.

Current SDK command:

```bash
bun run release:sdk:smoke
```

The previous audit-only path required consumers to list every internal `@ascend/*` package and add overrides. That path is now superseded for SDK validation by the checked-in smoke command above.

Current local SDK artifact shape:

1. Copy the built library package `dist` outputs into `/private/tmp/ascend-sdk-local-release/artifacts/packages`.
2. Bundle internal `@ascend/*` packages under the SDK artifact's own `node_modules`.
3. Pack a local SDK tarball.
4. Create a fresh consumer app whose `package.json` depends only on `@ascend/sdk` via that tarball.

This is still a local release-artifact smoke, not a final npm publishing design. It does prove the SDK can be installed and used without consumer overrides or repo-local package graph knowledge.

Implemented app path: `/private/tmp/ascend-apps-local-release`.

Current app command:

```bash
bun run release:apps:smoke
```

Current local app artifact shape:

1. Copy the built library and app `dist` outputs into `/private/tmp/ascend-apps-local-release/artifacts/packages`.
2. Bundle internal `@ascend/*` library packages under each app artifact's own `node_modules`.
3. Pack local CLI/API/MCP tarballs.
4. Create a fresh consumer app whose `package.json` depends only on those three app tarballs.

This proves installed app packages can run the workbook workflow without consumer overrides or repo-local package graph knowledge. It is still a local release-artifact smoke, not the final registry topology.

## Smokes

SDK smoke: passed with a single local SDK tarball and no consumer overrides.

Command:

```bash
bun run release:sdk:smoke
```

Workflow proven: create/open workbook, inspect sheets, plan a `setCells` operation, commit to output, reopen, verify structural check, and confirm formula recalculation. Result included `planWouldSucceed: true`, `reopenedValid: true`, `B1 = 125`, `C1 = 250`.

CLI smoke: passed from the installed `@ascend/cli` tarball with no TUI dependency.

Command:

```bash
bun run release:apps:smoke
```

Workflow proven: installed `node_modules/.bin/ascend --version`, `create`, setup `write`, `inspect`, `plan`, `commit`, `check`, `read`, and `docs`. Result included `planWouldSucceed: true`, `B1 = 125`, `C1 = 250`, and `docHits: 5`.

API smoke: passed from the installed `@ascend/api` tarball.

Command:

```bash
bun run release:apps:smoke
```

Workflow proven: imported `createApiFetch` and `createServer` from installed `@ascend/api`, then called `/write` for setup and `/inspect`, `/plan`, `/commit`, `/check`, and `/read` for the workflow. Result included `planWouldSucceed: true`, `B1 = 125`, `C1 = 250`, and `apiCapabilities: 66`.

MCP smoke: passed from the installed `@ascend/mcp` tarball.

Command:

```bash
bun run release:apps:smoke
```

Workflow proven: imported `createServer` from installed `@ascend/mcp`, confirmed `32` tools, then called `ascend.write`, `ascend.inspect`, `ascend.plan`, `ascend.commit`, `ascend.check`, `ascend.read`, and `ascend.search_docs`. Result included `planWouldSucceed: true`, `B1 = 125`, `C1 = 250`, `docHits: 5`, and `capabilities: 66`.

## Blocker Table

| Owner | File | Fix shape | Validation command | Release impact |
| --- | --- | --- | --- | --- |
| SDK/package build | `scripts/release-sdk-smoke.ts`, `package.json` | Done for local SDK smoke: produce a single local SDK tarball with bundled internal packages and verify a fresh consumer install without overrides. | `bun run release:sdk:smoke` | SDK external install path passes. |
| App package smokes | `scripts/release-apps-smoke.ts`, `package.json` | Done for local CLI/API/MCP smokes: produce local app tarballs with bundled internal packages and verify a fresh consumer install without overrides. | `bun run release:apps:smoke` | CLI/API/MCP external install workflow path passes. |
| CLI/headless | `apps/cli/package.json`, `apps/cli/src/index.ts`, `apps/cli/src/commands/tui.ts`, `apps/cli/src/commands/open.ts`, `apps/cli/tsconfig.json` | Done: CLI no longer depends on or project-references current TUI; `tui` and `open` lazy-load `@ascend/tui`; `create --json` is accepted by the central flag gate. | `bun test apps/cli/src/cli.test.ts`; `bun run release:apps:smoke` | Removes current TUI and CLI workflow flag blockers. |
| API build/startup | `scripts/build-packages.ts`, `apps/api/src/index.ts`, `apps/api/dist/package.json` | Done: real `dist/index.js`, generated publish manifest, root exports for `createApiFetch/createServer`, and `bin.ascend-api`; installed API workflow smoke passes. | `bun run release:apps:smoke` | Removes API dist/startup/packageability blocker. |
| MCP build/startup | `scripts/build-packages.ts`, `apps/mcp/src/index.ts`, `apps/mcp/dist/package.json` | Done: real `dist/index.js`, generated publish manifest, root export for `createServer`, and `bin.ascend-mcp`; installed MCP workflow smoke passes. | `bun run release:apps:smoke` | Removes MCP dist/startup/packageability blocker. |
| Bundled docs | `packages/sdk/src/agent-docs.ts`, `scripts/build-packages.ts`, `scripts/release-sdk-smoke.ts`, `scripts/release-apps-smoke.ts` | Done for SDK, CLI, and MCP installed surfaces. API does not expose a docs surface today. | `bun run release:sdk:smoke`; `bun run release:apps:smoke` | Removes agent-doc dependency on repo-local context for installed public docs surfaces. |
| Final publish topology | generated `dist/package.json` artifacts | Deferred by release owner: local artifact smokes use bundled internals; final registry/tarball-set policy is not decided in this phase. | N/A | Does not block current local external adoption proof; still required before public npm release. |

## Release Owner Next Step

Packageability implementation is now at the local external-adoption proof point. The remaining release-owner decision is whether bundled internal packages, a local tarball set, a local registry smoke, or a public npm package graph is the final publish topology.
