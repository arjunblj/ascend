# Release App Workflow Smoke

## Question

Can the local CLI/API/MCP app tarball smoke prove installed workflow behavior, not just package importability and capabilities callbacks?

## Hypothesis

Yes. The existing `release:apps:smoke` harness can exercise create/write/inspect/plan/commit/check/read through installed CLI, API, and MCP packages while keeping publication, provenance, production listener lifecycle, and real MCP stdio protocol claims blocked.

## External sources checked

- Bun package manager install docs: https://bun.sh/docs/cli/install
- npm package `bin` field: https://docs.npmjs.com/cli/v10/configuring-npm/package-json#bin
- Model Context Protocol specification: https://modelcontextprotocol.io/specification/2024-11-05/index
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

Release packageability is a supporting proof for the top claim handoffs. If app tarballs install but only expose capabilities callbacks, product owners still cannot tell whether the installed surfaces can perform a minimal workbook workflow. The North Star needs release proof that is useful without pretending local tarballs are registry publication or signed provenance.

## Probe/implementation

Folded the in-flight smoke harness into `scripts/release-apps-smoke.ts`:

- install CLI/API/MCP tarballs into a temp consumer;
- create temp workbooks;
- write setup operations;
- inspect, plan, commit with a hash guard, check, and read `B1:C1`;
- assert reopened values `B1=125` and `C1=250` through CLI, API, and MCP;
- assert installed CLI and MCP docs search for `plan commit`;
- keep capabilities checks for API and MCP.

Updated release packageability evidence in `fixtures/benchmarks/release-proof-index.ts` and the research claim board so the evidence wording matches the deeper smoke.

## Results

Validation passed:

```bash
bun run release:apps:smoke
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check scripts/release-apps-smoke.ts fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bunx tsc --build
bun run test:changed
```

The app smoke installed CLI/API/MCP tarballs into a temp consumer, ran the installed `ascend` bin, then proved CLI/API/MCP create/write/inspect/plan/commit/check/read workflows over temp workbooks. Each surface reopened `B1=125` and `C1=250`; CLI and MCP docs searches each returned 5 hits; API capabilities returned 66 capabilities; MCP registered 32 tools and returned 66 capabilities.

The claim boundary remains unchanged: this is local tarball install and installed workflow smoke only.

## Confidence

High for local installed workflow smoke. Medium for release packageability because the harness still avoids API listener lifecycle and MCP stdio protocol sessions.

## Fold-in decision

Promote to release proof support if validation passes. Do not promote a new product claim or surface.

## Next question

Should release owners require a real API listener lifecycle smoke and a real MCP stdio session before any app packageability wording moves beyond local tarball workflow evidence?
