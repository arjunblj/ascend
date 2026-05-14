# Agent Safe Edit: HTTP

Runnable transcript for coding agents using the local HTTP API. Start the server first:

```bash
bun run apps/api/src/index.ts
```

## Inspect And Read

```bash
curl -s http://localhost:3000/inspect \
  -H 'content-type: application/json' \
  -d '{"file":"model.xlsx"}'

curl -s http://localhost:3000/read \
  -H 'content-type: application/json' \
  -d '{"file":"model.xlsx","sheet":"Revenue","range":"A1:H50","format":"compact","rowLimit":50}'
```

Expected fields: `ok`, `data.sheets`, `data.compatibility`, compact `data.cells`, and `data.changeToken`.

## Repair Formula Text Before Planning

```bash
curl -s http://localhost:3000/formula-assist \
  -H 'content-type: application/json' \
  -d '{"formula":"=SUM(B2:G2","cursor":9,"prefix":"SU","functionName":"SUM","cycleReference":true}'
```

Expected fields: `data.diagnostics.parseOk`, `data.tokens`, `data.completions`, `data.signatureHelp`, and optional `data.cycle`.

## Plan With A Prepared Handle

```bash
curl -s http://localhost:3000/plan \
  -H 'content-type: application/json' \
  -d '{
    "file":"model.xlsx",
    "mutations":[
      { "path":"/sheets/Revenue/cells/H2/formula", "value":"=SUM(B2:G2)" }
    ],
    "compact":true,
    "prepare":true,
    "maxChangedCells":20
  }'
```

Expected fields: `data.inputSha256`, `data.planDigest`, `data.preview.wouldSucceed`, `data.approvals`, `data.lossAudit`, `data.preparedPlan.id`, and `data.modelOutput.nextActions`.

## Commit And Verify

```bash
curl -s http://localhost:3000/commit \
  -H 'content-type: application/json' \
  -d '{
    "planHandle":"<data.preparedPlan.id>",
    "output":"model.updated.xlsx",
    "compact":true,
    "maxAffectedCells":20
  }'

curl -s http://localhost:3000/check -H 'content-type: application/json' -d '{"file":"model.updated.xlsx"}'
curl -s http://localhost:3000/lint -H 'content-type: application/json' -d '{"file":"model.updated.xlsx"}'
curl -s http://localhost:3000/diff -H 'content-type: application/json' -d '{"fileA":"model.xlsx","fileB":"model.updated.xlsx"}'
```

If commit fails because the handle is expired, unavailable, or already used, re-run `plan`. If the workbook changed, re-read the affected range and re-plan with the new `inputSha256`.
