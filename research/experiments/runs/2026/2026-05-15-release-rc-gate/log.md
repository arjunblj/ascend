# Release RC Gate

## Question

Can Ascend prove an isolated RC-style packaged install across SDK, CLI, API, and MCP without relying on workspace imports or consumer package-manager overrides?

## Hypothesis

Yes. A release gate script can build JS artifacts, pack SDK/CLI/API/MCP tarballs, install only those tarballs into a temporary consumer app, and run a workbook proof across all four surfaces.

## External sources checked

- GitHub artifact attestation documentation: https://docs.github.com/en/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
- GitHub artifact attestation concepts: https://docs.github.com/en/actions/concepts/security/artifact-attestations
- SLSA provenance model: https://slsa.dev/spec/v1.0-rc1/provenance
- npm package `bundledDependencies` documentation: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#bundleddependencies
- Bun package manager documentation: https://bun.sh/docs/pm/cli/install

## Why this matters to Ascend

The release-proof bundle claim needs real packaged adoption evidence, not just source-tree tests. Agents and users should be able to install the published-style artifacts and run inspect/plan/commit/reopen/check/read workflows without workspace dependency leakage.

## Probe/implementation

Implemented `scripts/release-rc-gate.ts` and `bun run release:rc:gate`.

The gate:

- runs `bun run build:js`.
- packs `@ascend/sdk`, `@ascend/cli`, `@ascend/api`, and `@ascend/mcp` tarballs into `/private/tmp/ascend-rc-gate/artifacts`.
- installs only those tarballs into `/private/tmp/ascend-rc-gate/consumer`.
- rejects consumer overrides/resolutions.
- rejects installed Ascend manifests that retain `workspace:` or `file:` dependencies.
- runs a workbook proof across SDK, CLI, API fetch, and MCP registered tools.

## Results

First run failed under restricted network because the isolated consumer install needed public npm dependencies:

- `fast-xml-parser`
- `fflate`
- `zod`
- `@modelcontextprotocol/sdk`

After approved network access, the gate passed:

- SDK inspected `Sheet1`, reopened valid, and read `B1=125`, `C1=250`.
- CLI plan succeeded, check passed, read `B1=125`, `C1=250`, docs search returned 5 hits.
- API exported `createApiFetch` and `createServer`, committed the workbook, and reported 66 capabilities.
- MCP exported `createServer`, inspected `Sheet1`, check passed, read `B1=125`, `C1=250`, docs search returned 5 hits, and registered 32 tools.
- proof path: `/private/tmp/ascend-rc-gate/consumer/rc-gate-proof.json`.

Validation:

- `bun run release:rc:gate` failed without network as expected.
- `bun run release:rc:gate` passed with approved network access.
- `bunx biome check package.json scripts/release-rc-gate.ts`
- `bunx tsc --build`

## Confidence

Medium-high for local RC packageability. It proves local tarball install and runtime contracts across four surfaces, but it is not registry publication, signed provenance, artifact attestation, or retention/privacy policy.

## Fold-in decision

Promote to release loop as a harness. Do not promote to headline release proof until `release-proof-index` records the RC gate and release owners approve publication and non-attestation wording.

## Next question

Should `release-proof-index` consume `release:rc:gate` output as release packageability evidence, or should it wait until the remaining dirty SDK/IO/formula changes are committed and the worktree is cleaner?
