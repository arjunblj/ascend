# Agent Safe Edit: MCP

Use this transcript with any MCP client connected to `apps/mcp/src/index.ts`.

## Open Plan

```json
{ "tool": "ascend.open_plan", "arguments": { "file": "model.xlsx", "intent": "edit-plan" } }
```

Expected fields: `data.recommendedLoadOptions`, `data.reviewBeforeHydration`, `data.riskFeatures`, and `data.reasons`.

## Trust Preflight

```json
{ "tool": "ascend.trust_report", "arguments": { "file": "model.xlsx", "maxFindings": 50 } }
```

Expected fields: `data.trust`, `data.posture`, `data.includedInAgentContext`, `data.executionPolicy`, `data.findings[].code`, `data.findings[].location`, and `data.nextActions`.

## Inspect And Read

```json
{ "tool": "ascend.inspect", "arguments": { "file": "model.xlsx" } }
```

```json
{
  "tool": "ascend.read",
  "arguments": {
    "file": "model.xlsx",
    "sheet": "Revenue",
    "range": "A1:H50",
    "format": "compact",
    "rowLimit": 50
  }
}
```

Expected fields: `data.sheets`, `data.compatibility`, compact `data.cells`, and `data.changeToken`.

## Ask For Formula Help

```json
{
  "tool": "ascend.formula_assist",
  "arguments": {
    "formula": "=SUM(B2:G2",
    "cursor": 9,
    "prefix": "SU",
    "functionName": "SUM",
    "cycleReference": true
  }
}
```

Expected fields: `data.diagnostics.parseOk`, `data.tokens`, `data.completions`, `data.signatureHelp`, and optional `data.cycle`.

## Plan With A One-Shot Handle

```json
{
  "tool": "ascend.plan",
  "arguments": {
    "file": "model.xlsx",
    "mutations": [
      { "path": "/sheets/Revenue/cells/H2/formula", "value": "=SUM(B2:G2)" }
    ],
    "compact": true,
    "prepare": true,
    "maxChangedCells": 20
  }
}
```

Expected fields: `data.inputSha256`, `data.planDigest`, `data.preview.wouldSucceed`, `data.approvals`, `data.lossAudit`, `data.preparedPlan.id`, and `data.modelOutput.nextActions`.

## Commit And Verify

```json
{
  "tool": "ascend.commit",
  "arguments": {
    "planHandle": "<data.preparedPlan.id>",
    "output": "model.updated.xlsx",
    "compact": true,
    "maxAffectedCells": 20
  }
}
```

```json
{ "tool": "ascend.check", "arguments": { "file": "model.updated.xlsx" } }
```

```json
{ "tool": "ascend.lint", "arguments": { "file": "model.updated.xlsx" } }
```

```json
{ "tool": "ascend.diff", "arguments": { "fileA": "model.xlsx", "fileB": "model.updated.xlsx" } }
```

If `planHandle` is unavailable, expired, or already used, call `ascend.plan` again. Pass only exact approval ids emitted by the plan.
