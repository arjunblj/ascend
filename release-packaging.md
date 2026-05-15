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

## Verdict

Ascend is no longer blocked on the first SDK external install smoke, the current-TUI startup split, or the existence of real CLI/API/MCP dist JS and bins. The broader release remains blocked until packaged CLI/API/MCP external install smokes and installed-package docs are handled.

The SDK workflow now works from a temp external app with a single local `@ascend/sdk` tarball and no consumer overrides. CLI/API/MCP now build real app entrypoints and publish manifests, but they still need the same external packaged workflow smoke standard as SDK. Installed-package docs are still missing.

Previous audit stop condition was reached; current implementation sequence moves next to packaged docs/examples so agent docs work from installed packages.

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
- API exports `createApiFetch` and `createServer` from source, but no publish manifest exposes them from `dist`.
- MCP exports `createServer` from source, but no publish manifest or bin exposes it from `dist`.
- CLI bin points to `src/index.ts`, imports `./commands/tui.ts` at startup, and hard-depends on `@ascend/tui`.

Docs/OpenAPI:

- `bun test packages/sdk/src/agent-docs.test.ts` passed: OpenAPI currently lists implemented endpoints and docs/examples vocabulary is aligned inside the repo.
- External SDK package docs are not packaged. In the temp app, `readAgentDoc('llms.txt')` returned missing content and `searchAgentDocs({ query: 'plan commit' })` returned `0` hits.
- Root cause: `packages/sdk/src/agent-docs.ts` resolves docs from `new URL('../../../', import.meta.url)`, which works in repo layout but not in an installed `dist` package.

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

## Smokes

SDK smoke: passed with a single local SDK tarball and no consumer overrides.

Command:

```bash
bun run release:sdk:smoke
```

Workflow proven: create/open workbook, inspect sheets, plan a `setCells` operation, commit to output, reopen, verify structural check, and confirm formula recalculation. Result included `planWouldSucceed: true`, `reopenedValid: true`, `B1 = 125`, `C1 = 250`.

CLI smoke: workflow passed only after installing `@ascend/cli` from source and explicitly installing `@ascend/tui`.

Commands:

```bash
bun node_modules/.bin/ascend --version
bun node_modules/.bin/ascend inspect input.xlsx --json
bun node_modules/.bin/ascend plan input.xlsx --ops ops.json --json
bun node_modules/.bin/ascend commit input.xlsx --ops ops.json --output cli-output.xlsx --json --compact
bun node_modules/.bin/ascend check cli-output.xlsx --json
```

Result: version, inspect, plan, commit, and post-write check passed. Exact blocker: CLI is not packageable headlessly because `apps/cli/package.json` depends on `@ascend/tui`, and `apps/cli/src/index.ts` imports `./commands/tui.ts` at startup. This release should not require the current TUI.

API smoke: workflow passed only through installed source package deep import.

Command:

```bash
bun run api-smoke.ts
```

Workflow used `createApiFetch` from `@ascend/api/src/server.ts`, then called `/health`, `/inspect`, `/plan`, `/commit`, and `/check`. Result included `health.status: ok`, `planWouldSucceed: true`, `postWrite.valid: true`, and `checkValid: true`.

Exact blocker: no packageable API `dist` manifest, no bin, and app JS build output is placeholder.

MCP smoke: workflow passed only through installed source package deep import/internal registered-tool access.

Command:

```bash
bun run mcp-smoke.ts
```

Workflow used `createServer` from `@ascend/mcp/src/index.ts`, confirmed `32` tools, then called `ascend.inspect`, `ascend.plan`, `ascend.commit`, and `ascend.check`. Result included `planWouldSucceed: true`, `postWrite.valid: true`, and `checkValid: true`.

Exact blocker: no packageable MCP `dist` manifest, no bin, and app JS build output is placeholder.

## Blocker Table

| Owner | File | Fix shape | Validation command | Release impact |
| --- | --- | --- | --- | --- |
| SDK/package build | `scripts/release-sdk-smoke.ts`, `package.json` | Done for local SDK smoke: produce a single local SDK tarball with bundled internal packages and verify a fresh consumer install without overrides. Next step is deciding whether this bundled shape or a registry/tarball-set shape is the final publish design. | `bun run release:sdk:smoke` | First SDK external install smoke passes; final publish design still needs release decision. |
| Package manifests | `packages/*/package.json` | Decide source manifests vs generated publish manifests; source manifests should not remain the only package truth with `private: true`, `main: src/index.ts`, no `exports`, and `workspace:*`. | `bun run build` plus manifest audit showing publishable manifests for every public package. | Blocks publish readiness and consumer trust. |
| CLI/headless | `apps/cli/package.json`, `apps/cli/src/commands/tui.ts`, `apps/cli/src/commands/open.ts`, `apps/cli/tsconfig.json` | Done for source startup: CLI no longer depends on or project-references current TUI; `tui` and `open` lazy-load `@ascend/tui` only when invoked. Still needs external packaged CLI smoke after real dist/bin exists. | `bun test apps/cli/src/cli.test.ts`; `bun run build` | Removes current TUI as a headless CLI startup blocker. |
| CLI build/bin | `scripts/build-packages.ts`, `apps/cli/dist/package.json` | Done for built artifact shape: real `dist/index.js`, generated publish manifest, and `bin.ascend`. Still needs external install smoke using produced app artifacts. | `bun apps/cli/dist/index.js --version`; manifest audit | Removes CLI dist/bin blocker. |
| API build/startup | `scripts/build-packages.ts`, `apps/api/src/index.ts`, `apps/api/dist/package.json` | Done for built artifact shape: real `dist/index.js`, generated publish manifest, root exports for `createApiFetch/createServer`, and `bin.ascend-api`. Still needs external install smoke. | import `apps/api/dist/index.js`; call `/health` through `createApiFetch` | Removes API dist/startup blocker. |
| MCP build/startup | `scripts/build-packages.ts`, `apps/mcp/src/index.ts`, `apps/mcp/dist/package.json` | Done for built artifact shape: real `dist/index.js`, generated publish manifest, root export for `createServer`, and `bin.ascend-mcp`. Still needs external install smoke. | import `apps/mcp/dist/index.js`; verify `32` registered tools | Removes MCP dist/startup blocker. |
| Bundled docs | `packages/sdk/src/agent-docs.ts`, `docs/*`, `examples/*`, `llms*.txt` | Package docs/examples as assets or embed them at build time; resolve relative to package, not repo root. | External temp app: `readAgentDoc('llms.txt')` returns text and `searchAgentDocs({ query: 'plan commit' })` returns workflow hits. | Blocks agent use without repo-local context. |
| App dist hygiene | `apps/*/dist/index.js`, app build config | Remove stale placeholder JS; app build must generate runtime JS or omit misleading files. | `bun run build` then inspect app `dist/index.js`; startup smokes use only built app files. | Blocks confidence in release artifacts. |

## Release Owner Next Step

Implement packageability in this order:

1. Package docs/examples with SDK/API/MCP so agents do not need the repo checkout.
2. Add external packaged CLI/API/MCP smokes using the same artifact standard as `release:sdk:smoke`.
