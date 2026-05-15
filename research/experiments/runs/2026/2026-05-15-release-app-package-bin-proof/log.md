# Release App Package Bin Proof

Date: 2026-05-15

## Question

Can Ascend's release packaging proof move past SDK-only smoke by building CLI, API, and MCP app artifacts with publishable bin manifests, bundled SDK agent docs, and side-effect-free imports?

## Hypothesis

If app entrypoints use an explicit direct-run guard, the build script emits app dist manifests with `bin` entries, and SDK agent docs are copied into the release artifact, then packageability can be improved without adding any new SDK, CLI, API, or MCP product surface.

## External sources checked

- Bun bundler documentation: https://bun.sh/docs/bundler
- npm `package.json` `bin` documentation: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#bin
- Node.js ESM URL utilities and `pathToFileURL`: https://nodejs.org/api/url.html#urlpathtofileurlpath-options
- Node.js ESM import expressions: https://nodejs.org/api/esm.html#import-expressions

## Why this matters to Ascend

The ranked portfolio keeps "release proof bundle" below the top two product claims because packaged adoption is not yet proven. A headless external consumer must be able to install artifacts and import entrypoints without accidentally starting servers, MCP transports, or TUI flows.

## Probe/implementation

- `scripts/build-packages.ts` now builds `apps/cli`, `apps/api`, and `apps/mcp` into `dist/index.js` using `Bun.build` with external packages.
- App dist manifests now include bin entries: `ascend`, `ascend-api`, and `ascend-mcp`.
- `apps/cli/src/index.ts`, `apps/api/src/index.ts`, and `apps/mcp/src/index.ts` now use `pathToFileURL(process.argv[1]).href` direct-run guards instead of import-time side effects.
- `apps/api/src/index.ts` re-exports API server helpers from `server.ts` so importing the package does not start a listener.
- SDK agent documentation assets are copied into the SDK dist package so runtime docs do not depend on the repo layout.
- `WorkbookDocument.writePlanSummary()` gives partial/full document sessions the same write-plan guard semantics used by agent workflow code.

## Results

Proof commands run:

```bash
bun run build:js
bun -e "await import('./apps/cli/src/index.ts'); await import('./apps/api/src/index.ts'); await import('./apps/mcp/src/index.ts'); console.log('app entrypoint imports did not auto-run')"
bun -e "const fs=await import('node:fs/promises'); for (const p of ['apps/cli/dist/package.json','apps/api/dist/package.json','apps/mcp/dist/package.json']) { const j=JSON.parse(await fs.readFile(p,'utf8')); console.log(p, JSON.stringify({name:j.name,main:j.main,bin:j.bin,dependencies:j.dependencies})) }"
bun run release:sdk:smoke
bun test apps/cli/src/cli.test.ts -t "--version prints version through the executable boundary|plan and commit implement safe agent workflow"
bun test apps/cli/src/cli.test.ts -t "docs --json searches bundled agent docs"
bun test apps/api/src/server.test.ts -t "prepared plan handles commit without reopening operation input"
bun test apps/mcp/src/index.test.ts -t "ascend.commit accepts prepared plan handles"
bun test packages/sdk/src/sdk.test.ts -t "WorkbookDocument exposes read-side model and write summary for cached verification"
```

Observed evidence:

- `build:js` emitted library and app JS artifacts.
- Importing CLI/API/MCP source entrypoints printed `app entrypoint imports did not auto-run`.
- Dist manifests emitted `bin` entries for all three apps and rewrote internal dependencies to package versions.
- SDK dist includes the agent docs needed by packaged consumers.
- SDK release smoke still passed through create/open/plan/commit/reopen/check/recalc.
- CLI, API, MCP, and SDK targeted behavior tests passed.

## Confidence

Medium-high for "app entrypoint artifacts, bin manifests, packaged SDK docs, and import-safe entrypoints are buildable." This is not yet a full app tarball consumer proof because the smoke still installs only the SDK tarball.

## Fold-in decision

Promote to product/release loop as a small implementation proof already folded into code. Keep release claims conservative: this closes part of the release packageability blocker, not publication policy, signed artifacts, app tarball installation, docs packaging, or provenance.

## Next question

Can a release owner extend the smoke harness to pack and install CLI/API/MCP app tarballs in a temp consumer and prove `ascend --version`, API import, and MCP import without workspace dependencies?
