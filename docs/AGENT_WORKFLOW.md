# Ascend Agent Workflow

Use this workflow for headless spreadsheet edits:

1. Inspect: `ascend inspect <file> --json --verbose`
2. Locate: `ascend read`, `ascend read table`, `ascend find`, and `ascend inspect --detail`
3. Build operations from `ascend ops --json` or `ascend://operations`
4. Plan: `ascend plan <file> --ops ops.json --progress jsonl --json`
5. Commit: `ascend commit <file> --ops ops.json --output out.xlsx --expect-sha256 <hash> --progress jsonl --json`
6. Verify: `ascend check`, `ascend lint`, `ascend diff`, `ascend trace`, or `ascend export`
7. Recover: `ascend repair-plan <file> --json`

Safety defaults:

- Prefer `--output` over `--in-place`.
- Use the plan `inputSha256` as `--expect-sha256`.
- Use only approval ids emitted by plan.
- Use `--allow-loss` only for explicit user-approved feature loss.
- Keep stdout machine-readable and consume `--progress jsonl` from stderr.

MCP resources:

- `ascend://llms.txt`
- `ascend://llms-full.txt`
- `ascend://docs/agent-api.md`
- `ascend://capabilities`
- `ascend://operations`
- `ascend://agent-workflow`

Recovery search:

- `ascend.search_docs` for commands, workflow guidance, schemas, and safety policy.
- `ascend.search_examples` for runnable SDK examples and MCP setup snippets.
