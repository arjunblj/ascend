# Release App Tarball Smoke

Date: 2026-05-15

## Question

Can Ascend prove that CLI, API, and MCP app release artifacts install into a temp external consumer without workspace dependencies?

## Hypothesis

If app tarballs bundle Ascend internal packages and keep only public dependencies in their publish manifests, then a consumer can install CLI/API/MCP artifacts from tarballs, run the `ascend` bin, and import API/MCP modules outside the monorepo.

## External sources checked

- Bun package manager pack workflow: https://bun.sh/docs/cli/pm
- Bun package manager install workflow: https://bun.sh/docs/cli/install
- npm `package.json` `bin` documentation: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#bin
- npm bundled dependencies documentation: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#bundleddependencies

## Why this matters to Ascend

The release proof bundle was blocked by "works only in the workspace" evidence. A product-shaped release claim needs a consumer proof for the actual artifacts agents and humans would install, not only source-tree tests.

## Probe/implementation

- Added `scripts/release-apps-smoke.ts`.
- Added `release:apps:smoke` to the root scripts.
- The harness builds JS artifacts, stages package and app dist folders under `/private/tmp/ascend-apps-local-release`, bundles internal `@ascend/*` package artifacts into each app tarball, packs `@ascend/cli`, `@ascend/api`, and `@ascend/mcp`, installs only those tarballs into a temp consumer, runs `node_modules/.bin/ascend --version`, and imports API/MCP exports from the installed packages.

## Results

Proof commands run:

```bash
bun run release:apps:smoke
bunx biome check package.json scripts/release-apps-smoke.ts
bunx tsc --build
```

Observed output:

- `ascend --version` from the temp consumer printed `0.0.0`.
- API import reported `createApiFetch` and `createServer` as functions.
- MCP import reported `createServer` as a function.
- App artifacts were produced at `/private/tmp/ascend-apps-local-release/artifacts/ascend-cli-0.0.0.tgz`, `ascend-api-0.0.0.tgz`, and `ascend-mcp-0.0.0.tgz`.

## Confidence

Medium-high for local app tarball installation and import safety. This does not prove registry publication, signed artifacts, provenance, production API server lifecycle, or an MCP protocol session.

## Fold-in decision

Promote to release/product loop as a reusable smoke harness. This is a proof harness, not a new product surface. It narrows the release proof bundle blocker from "app artifacts are not installable" to "publication policy, artifact signing/provenance, and deeper protocol/runtime smoke remain owner-owned."

## Next question

Should the release owner extend the app smoke to execute one API request through `createApiFetch` and one MCP tool registration/read path, or keep those as separate product/protocol smoke gates?
