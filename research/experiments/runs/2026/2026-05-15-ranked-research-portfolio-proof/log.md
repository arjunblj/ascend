# Ranked Research Portfolio Proof

## Question

Can Ascend stop broad topic sweeping and manage research as a ranked portfolio with proof-producing gates for the top unknowns?

## Hypothesis

Yes. A portfolio with kill criteria should reduce surface sprawl. The top unknowns should be proven with existing harnesses and tests, not new product surfaces.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- LSP 3.17: https://microsoft.github.io/language-server-protocol/specifications/lsp/3.17/specification/
- Excel structured references: https://support.microsoft.com/en-gb/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
- Model Context Protocol: https://modelcontextprotocol.io/specification/2024-11-05/index
- Univer MCP guide: https://docs.univer.ai/guides/sheets/getting-started/mcp
- PostgreSQL MVCC: https://www.postgresql.org/docs/17/mvcc-intro.html
- SQLite isolation: https://www.sqlite.org/isolation.html
- RocksDB snapshots: https://github.com/facebook/rocksdb/wiki/Snapshot
- Automerge concepts: https://automerge.org/docs/reference/concepts/
- DuckDB Excel import: https://duckdb.org/docs/stable/guides/file_formats/excel_import
- Apache Arrow columnar format: https://arrow.apache.org/docs/format/Columnar.html

## Why this matters to Ascend

The North Star needs evidence-backed bets, not a growing list of interesting surfaces. Ranking directions by claim, evidence, kill criterion, and owner loop makes research accountable to correctness, performance, product/DX, or release proof ownership.

## Probe/implementation

- Inspected the release claim board, claim ladder, agent-context synthesis, and current proof harnesses.
- Finished a small in-flight shared-formula grouping fix separately:
  - `38b437e1 fix(engine): scope shared formula groups by master`
  - `de9d4ac9 test(engine): cover shared formula group collisions`
- Produced proof for top unknown 1: formula refusal across CLI/API/MCP:
  - `be666996 test(sdk): prove formula rename refusals across surfaces`
- Produced proof for top unknown 2: retained viewport patch history:
  - reran viewport proof harness and SDK/API/MCP validations.
- Added `research/experiments/syntheses/2026-05-ranked-research-portfolio.md`.

## Results

Portfolio directions ranked:

1. Safe unknown workbook opening.
2. Auditable package-part mutation.
3. Formula rejection-first language service.
4. Retained viewport patch history.
5. Token-bounded agent view.
6. Release proof index.
7. Formula conformance/oracle routing.
8. Property-based journal laws.
9. Columnar scan sidecars.
10. Agent workflow observability.

Top proof result 1: formula assist now has cross-surface table-name rename refusal snapshots. Rename remains killed.

Top proof result 2: viewport retained patch history passed the tracked harness and SDK/API/MCP validations. CLI remains excluded; CRDT/collaboration language remains killed.

## Confidence

High that the portfolio is now narrow enough to steer work. High that the two selected unknowns received proof rather than prose. Medium on rank ordering after item 6 because property-based journal laws and columnar sidecars may trade places after fresh benchmark/law evidence.

## Fold-in decision

Promote to topic synthesis and proof evidence only. Do not add new product surfaces. Hand off formula latency/corpus proof and viewport product example as the next focused proof moves.

## Next question

Can formula language-service latency and rejection behavior be proven over a public formula corpus without adding rename or workbook-context edits?
