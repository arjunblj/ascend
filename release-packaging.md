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

## Verdict

Ascend is no longer blocked on the first SDK external install smoke. The broader release remains blocked until CLI, API, MCP, and installed-package docs are handled.

The SDK workflow now works from a temp external app with a single local `@ascend/sdk` tarball and no consumer overrides. CLI, API, and MCP workflows still work only through source-package installs/deep imports or the current TUI dependency. That is not a complete external adoption path.

Previous audit stop condition was reached; current implementation sequence moves next to the headless CLI/TUI split.

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
| CLI/headless | `apps/cli/package.json`, `apps/cli/src/index.ts` | Remove hard runtime dependency on current `@ascend/tui`; make TUI optional/dynamic or move it to a separate package. Ship headless CLI commands without requiring TUI. | Temp app install `@ascend/cli`; run `ascend --version`, `inspect`, `plan`, `commit`, `check` with no `@ascend/tui` dependency. | Blocks headless agent workflow release. |
| CLI build/bin | `apps/cli/tsconfig.json`, `apps/cli/package.json`, build script | Produce real `dist` JS and a publish manifest with `bin.ascend` pointing at built JS, not `src/index.ts`. | `node_modules/.bin/ascend --version` from installed artifact, without source files. | Blocks CLI packageability. |
| API build/startup | `apps/api/tsconfig.json`, `apps/api/package.json`, `apps/api/src/index.ts` | Produce real `dist` JS, publish manifest, explicit exports for `createApiFetch/createServer`, and a server bin such as `ascend-api`. | Temp app install `@ascend/api`; import package root and run server/bin smoke for `/health`, `/plan`, `/commit`, `/check`. | Blocks HTTP API release. |
| MCP build/startup | `apps/mcp/tsconfig.json`, `apps/mcp/package.json`, `apps/mcp/src/index.ts` | Produce real `dist` JS, publish manifest, root export for `createServer`, and MCP server bin for stdio startup. | Temp app install `@ascend/mcp`; run MCP stdio/server smoke for `ascend.inspect`, `ascend.plan`, `ascend.commit`, `ascend.check`. | Blocks MCP release. |
| Bundled docs | `packages/sdk/src/agent-docs.ts`, `docs/*`, `examples/*`, `llms*.txt` | Package docs/examples as assets or embed them at build time; resolve relative to package, not repo root. | External temp app: `readAgentDoc('llms.txt')` returns text and `searchAgentDocs({ query: 'plan commit' })` returns workflow hits. | Blocks agent use without repo-local context. |
| App dist hygiene | `apps/*/dist/index.js`, app build config | Remove stale placeholder JS; app build must generate runtime JS or omit misleading files. | `bun run build` then inspect app `dist/index.js`; startup smokes use only built app files. | Blocks confidence in release artifacts. |

## Release Owner Next Step

Implement packageability in this order:

1. Split current TUI out of the CLI runtime path for this release.
2. Produce real CLI/API/MCP `dist` JS plus publish manifests and bins.
3. Package docs/examples with SDK/API/MCP so agents do not need the repo checkout.
