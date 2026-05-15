# Open Plan API Endpoint

## Question

Can the HTTP API expose the open planner without encouraging clients to skip existing trust reports and package-graph audits?

## Hypothesis

Yes. A standalone `/open-plan` endpoint can return the same SDK schema as CLI/MCP and stay narrowly scoped to load-mode recommendation, while `/trust-report` and `/package-graph` remain the detailed risk and evidence endpoints.

## External sources checked

- Microsoft Graph workbook `createSession`, which makes session behavior explicit through a request body and response object: https://learn.microsoft.com/en-us/graph/api/workbook-createsession
- Google Sheets `spreadsheets.get`, which recommends field masks and limited grid data for large spreadsheets: https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/get
- SheetJS parse options for explicit caller-selected workbook metadata, sheet, formula, raw-file, and VBA reads: https://docs.sheetjs.com/docs/api/parse-options/
- openpyxl optimized read-only mode for large XLSX files: https://openpyxl.pages.heptapod.net/openpyxl/optimized.html
- Postman REST API best-practices guide on predictable endpoint naming: https://blog.postman.com/rest-api-best-practices/
- Model Context Protocol schema/tool docs as a parallel schema-first agent API reference: https://modelcontextprotocol.io/specification/2025-06-18/schema

## Why this matters to Ascend

API clients need the same cheap pre-hydration planning primitive as CLI and MCP clients. Keeping it standalone prevents implicit behavior changes in `/inspect` and gives service clients a predictable way to route workbooks before calling heavier endpoints.

## Probe/implementation

Local inspection:

- `apps/api/src/server.ts` already groups lightweight endpoints before mutation endpoints.
- `/inspect`, `/active-content`, `/trust-report`, `/package-graph`, and `/raw-part` have separate responsibilities.
- `apps/api/api.test.ts` has compact endpoint tests using `createApiFetch`.

Implementation:

- Added `POST /open-plan`.
- Request body: `{ file, intent? }`.
- Valid intents: `risk-inventory`, `read-values`, `formula-analysis`, `edit-plan`.
- Implementation reads file bytes directly and calls `inspectWorkbookOpenPlan`, avoiding workbook hydration.
- Invalid intents return a structured `VALIDATION_ERROR` with allowed values.
- Added the endpoint to `docs/openapi.yaml` after the bundled-doc contract test caught the missing path.
- Existing `/inspect`, `/trust-report`, and `/package-graph` behavior is unchanged.

Validation:

- `bun test apps/api/api.test.ts -t "open-plan"` passed.
- `bun test packages/sdk/src/agent-docs.test.ts -t "OpenAPI lists"` passed.
- `bunx biome check apps/api/src/server.ts apps/api/api.test.ts` passed after formatting.
- `bunx tsc --build` passed.
- `bun run test:changed` initially failed because `docs/openapi.yaml` did not list `/open-plan`; after adding the OpenAPI entry, it passed: 4959 pass, 1 skip, 0 fail.

## Results

The API fold-in is useful and low-risk:

- Value-read intent returns `{ mode: "values" }`.
- Macro workbooks return `{ mode: "metadata-only" }` with `reviewBeforeHydration: true`.
- Bad intents are rejected before file reads.
- The endpoint complements, rather than replaces, `/trust-report` and `/package-graph`.

## Confidence

High that `/open-plan` should stay. Medium that clients will naturally call it first; future docs or agent-init guidance should mention it explicitly.

## Fold-in decision

Folded into production as `POST /open-plan`. Promote to agent workflow docs after one pass over `agent-init`, bundled docs, and API doc search snippets.

## Next question

Should `agent-init` and bundled agent docs teach the open-plan -> inspect/read/agent-view -> plan/commit route as the default unknown-workbook workflow?
