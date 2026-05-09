# Ascend Agent API

This document is the compact, markdown-native reference for agents using Ascend without browsing.

## CLI Contract

Use `ascend --help` for the command list and `ascend <command> --help` for command-specific flags.

Core commands:

- `ascend agent-init --json` prints the recommended machine workflow.
- `ascend ops --json` lists all operation schemas, examples, invalid examples, recovery actions, and approval metadata.
- `ascend capabilities --json` returns Excel feature coverage, priorities, OSS baseline notes, tests, gap reasons, and next milestones.
- `ascend inspect <file> --json --verbose` opens workbook metadata and compatibility context.
- `ascend read <file> <selector> --json` reads ranges, tables, named ranges, cells, rows, objects, compact cells, or TSV depending on flags.
- `ascend find <file> <query> --json` searches values and formulas.
- `ascend plan <file> --ops ops.json --progress jsonl --json` validates operations, previews diffs, audits recalc, approvals, and preservation risk.
- `ascend commit <file> --ops ops.json --output out.xlsx --expect-sha256 <hash> --progress jsonl --json` writes safely.
- `ascend repair-plan <file> --json` suggests recovery actions after failed checks, lints, or unsupported-feature audits.
- `ascend docs <query> --json`, `ascend docs --examples <query> --json`, `ascend docs --path llms-full.txt`, and `ascend docs --list --json` expose local agent docs without browsing.
- `ascend check <file> --json`, `ascend lint <file> --json`, `ascend trace <file> <ref> --json`, and `ascend diff <a> <b> --json` verify work.

JSON stdout is wrapped in a versioned machine envelope. Long-running safe workflows can emit JSONL progress to stderr with `--progress jsonl`.

## MCP Contract

Use these resources for stable context:

- `ascend://llms.txt` short agent map.
- `ascend://llms-full.txt` expanded agent context.
- `ascend://docs/agent-api.md` this reference.
- `ascend://operations` canonical operation catalog.
- `ascend://capabilities` Excel capability registry.
- `ascend://agent-workflow` safe workflow guide.

Use these discovery tools when stuck:

- `ascend.search_docs({ query, limit?, tokens? })` searches docs, references, workflow guidance, and release-facing agent context.
- `ascend.search_examples({ query, limit?, tokens? })` searches examples and MCP setup snippets.
- `ascend.list_operations()` returns canonical operation schemas and examples.
- `ascend.capabilities({ feature?, family?, priority?, status?, gapsOnly? })` returns capability coverage.

Use these workbook tools for normal work:

- `ascend.inspect({ file, sheet? })`
- `ascend.list_sheets({ file })`
- `ascend.read({ file, sheet?, range, format?, rowOffset?, rowLimit?, display?, headers? })`
- `ascend.read_table({ file, table, rowOffset?, rowLimit?, display? })`
- `ascend.find({ file, query, sheet?, in?, caseSensitive?, limit? })`
- `ascend.visuals({ file })`
- `ascend.agent_view({ file, sheet?, range, rowChunkSize?, sampleRowLimit?, sampleValueLimit? })`
- `ascend.plan({ file, ops })`
- `ascend.commit({ file, ops, output?, inPlace?, backup?, expectSha256?, allowLoss?, approvals? })`
- `ascend.repair_plan({ file })`
- `ascend.check({ file })`, `ascend.lint({ file })`, `ascend.trace({ file, cell })`, `ascend.diff({ fileA, fileB })`, `ascend.export({ file, output, format? })`

## Safety Rules

- Prefer non-destructive output paths over in-place edits.
- Use `inputSha256` from plan as `expectSha256` during commit.
- Pass only approval IDs emitted by plan.
- Pass `allowLoss` only when the user explicitly accepts the exact feature/tier loss.
- Treat macros, signatures, ActiveX, form controls, Power Query, data models, pivots, slicers, chartsheets, and other preserve-first features as high-risk unless plan says the write is safe.

## Operation Pattern

Create `ops.json` as an array:

```json
[
  {
    "op": "setCells",
    "sheet": "Sheet1",
    "updates": [{ "ref": "B2", "value": 42 }]
  }
]
```

Then:

```bash
ascend plan model.xlsx --ops ops.json --progress jsonl --json
ascend commit model.xlsx --ops ops.json --output model.updated.xlsx --expect-sha256 <inputSha256> --progress jsonl --json
ascend check model.updated.xlsx --json
```

## Example Recovery Prompts

- Search operation schema: `ascend.search_docs({ "query": "setChartSeriesSource chart series source" })`
- Find example code: `ascend.search_examples({ "query": "read modify save workbook" })`
- Understand safety gates: `ascend.search_docs({ "query": "allowLoss approvals preservation audit" })`
- Read compactly: `ascend.search_docs({ "query": "compact read formats MCP" })`
