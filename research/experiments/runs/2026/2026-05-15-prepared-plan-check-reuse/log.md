# Prepared Plan Check Reuse

## Question

Can prepared agent-plan commits reuse the structural check already computed during planning without hiding stale-input or formula-context risk?

## Hypothesis

Yes, but only as a narrowly invalidated prepared-plan optimization: reuse is credible for one-shot prepared commits when the input hash still matches, the commit uses literal `setCells` operations, and the prepared workbook preimage has no formula cells or formula metadata.

## External sources checked

- [PostgreSQL PREPARE](https://www.postgresql.org/docs/current/sql-prepare.html) separates parse/rewrite work from later execution and documents forced re-analysis/re-planning when dependent database objects change.
- [SQLite prepare APIs](https://sqlite.org/c3ref/prepare.html) frame prepared statements as compiled objects and document automatic recompile on schema changes or parameter values that can change a query plan.
- [SLSA software attestations](https://slsa.dev/spec/v1.1/attestation-model) reinforces that release claims should bind evidence to artifacts and avoid implying stronger provenance than the generated metadata provides.

## Why this matters to Ascend

Prepared agent plans are one of Ascend's strongest agent-DX surfaces: plan, inspect, then commit without reopening and recomputing every proof. If the commit can reuse a validated planning proof under explicit invalidation boundaries, the product claim becomes stronger: "trustworthy prepared workbook mutation" can be backed by one-shot handles, hash guards, post-write audits, and observable skipped work.

## Probe/implementation

- Added an internal `preparedCheck` channel from `createPreparedAgentPlan` into `commitAgentPlanFromWorkbook`.
- Added the same check handoff for prepared path-mutation handles in API and MCP flows.
- Reused the prepared check only when:
  - a prepared check exists,
  - every operation is `setCells`,
  - the prepared write-policy workbook has zero formula cells and zero formula-info cells.
- Added focused assertions that prepared SDK and MCP commits report `writePolicyCheckMs === 0`.
- Initial probe failed because a formula-free `setCells` apply still marks `recalcRequired: true`; the gate was corrected to key off the stronger formula-free preimage invariant instead of the engine's conservative recalc flag.

## Results

- `bun test packages/sdk/src/agent-workflow.test.ts -t "prepared agent plans reuse full workflow state|prepared agent plans expose rollback journal safety facts|prepared agent commits"` passed: 4 tests.
- `bun test apps/mcp/src/index.test.ts -t "prepared plan|planHandle|writePolicyCheckMs"` passed: 1 test.
- `bun test apps/api/src/server.test.ts -t "prepared|plan and commit"` passed: 18 tests.

The prepared SDK and MCP value-edit commits now prove the intended optimization by emitting a zero write-policy-check timing while retaining hash staleness guards and post-write reopen audits.

## Confidence

Medium. The reuse gate is deliberately narrow and test-backed for SDK/MCP/API prepared handles. It is not a general workflow cache and does not cover formula edits, formula-bearing workbooks, structural workbook edits, or multi-writer collaboration.

## Fold-in decision

Promote to product/DX loop and correctness loop as a tiny implementation proof. The allowed product wording is: "prepared value-edit commits can reuse the planning structural check under a formula-free, one-shot, hash-guarded contract."

Do not promote broader cache, collaboration, or formula-workbook wording.

## Next question

Should the release-claim board add a distinct "prepared agent commit proof" row, or should this remain supporting evidence under trustworthy mutation planning until larger real-workbook fixtures prove user-visible latency impact?
