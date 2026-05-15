# Release Packaging Owner Audit

## Question

Can the untracked release packaging audit become owner-routed proof without changing package surfaces prematurely?

## Hypothesis

Ascend's release proof bundle claim should stay blocked until packageability is proven from clean external installs. A local manifest/build audit can identify owner actions, but it should not promote SDK, CLI, API, or MCP release wording by itself.

## External sources checked

- npm `package.json` `files` field and publish packaging rules: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#files
- Node.js package `exports`: https://nodejs.org/api/packages.html#exports
- Bun build documentation: https://bun.sh/docs/bundler
- Bun single executable documentation: https://bun.sh/docs/bundler/executables
- Model Context Protocol specification: https://modelcontextprotocol.io/specification/2024-11-05/index

## Why this matters to Ascend

The North Star includes world-class agent DX and trustworthy release proof. A proof bundle that cannot be installed by external users without repo-local assumptions is not credible release evidence, even if local SDK, CLI, API, and MCP smokes pass in source form.

## Probe/implementation

Inspected the existing untracked `release-packaging.md` audit and verified its main blockers against current local files:

- `packages/*/package.json` source manifests are still `private: true`, point `main` at `src/index.ts`, and use `workspace:*` internal dependencies.
- `scripts/build-packages.ts` builds only library packages and rewrites generated `dist/package.json` manifests for `packages/schema`, `core`, `formulas`, `engine`, `io-xlsx`, `io-csv`, `verify`, and `sdk`.
- `apps/cli`, `apps/api`, and `apps/mcp` `tsconfig.json` files still use `emitDeclarationOnly: true`.
- App dist JS entrypoints are placeholders:
  - `apps/cli/dist/index.js`: `console.log("ascend cli")`
  - `apps/api/dist/index.js`: `console.log("ascend api")`
  - `apps/mcp/dist/index.js`: `console.log("ascend mcp")`
- `apps/cli/package.json` still points `bin.ascend` to `src/index.ts` and depends on `@ascend/tui`.
- API and MCP expose useful source functions, but no packageable dist manifest or bin currently owns them.
- `packages/sdk/src/agent-docs.ts` resolves docs from `new URL('../../../', import.meta.url)`, which depends on repo layout rather than packaged assets.

No production changes were made. The root `release-packaging.md` file was left untracked.

## Results

Owner-routed blocker table:

| Owner | Blocking proof gap | Evidence checked | Required next validation |
| --- | --- | --- | --- |
| SDK/package build | Clean external install requires local dist package graph or registry/tarball flow, not consumer overrides. | Library `dist/package.json` rewrites internal deps to `0.0.0`; source manifests are private/workspace-only. | Fresh temp app installs produced artifacts with no overrides. |
| CLI/headless | CLI package is not a headless publishable artifact. | `bin.ascend` points at `src/index.ts`; source imports TUI command at startup; app dist JS is placeholder. | Installed `@ascend/cli` runs `ascend --version`, `inspect`, `plan`, `commit`, and `check` using built JS only. |
| API build/startup | API has source server functions but no publishable dist manifest/bin. | `apps/api` emits declarations only; dist JS is placeholder. | Installed `@ascend/api` imports package root and runs `/health`, `/plan`, `/commit`, `/check`. |
| MCP build/startup | MCP has source server functions but no publishable dist manifest/bin. | `apps/mcp` emits declarations only; dist JS is placeholder. | Installed `@ascend/mcp` starts the MCP server and runs core spreadsheet tools. |
| Bundled docs | Agent docs are repo-layout dependent. | `agent-docs.ts` resolves from repo root via `import.meta.url`. | External installed SDK can read `llms.txt` and search `plan commit` without repo checkout. |

Decision: move this evidence into research ownership and keep the release proof bundle claim as `needs-one-more-fold-in`. Do not implement package surfaces in this research block.

## Confidence

High on the manifest/build/doc blockers because they are visible in source and dist files. Medium on the exact temporary install behavior because this cycle did not rerun the full external temp-app smoke.

## Fold-in decision

Promote to topic synthesis and release/product owner prompts only. Archive root `release-packaging.md` as untracked external evidence for now; do not commit it as a root doc.

## Next question

Should a future release-owner loop add a checked-in packaging smoke harness that builds local tarballs and installs SDK, CLI, API, and MCP into a temp app with no consumer overrides?
