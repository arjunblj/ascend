# 2026-05 Leapfrog Research Synthesis

Date: 2026-05-14

## Strongest Candidates

1. Package action proof is the clearest leapfrog claim. Current write plans already expose generated and copied-through parts, package graph audits, write policies, and trace digests. Fold this into a `PackageDeltaPlan` with actions: passthrough, regenerate, add, drop, error. This supports both correctness and dirty-write performance.
2. Release proof bundle is the clearest product/DX claim. The probe produced real inspect, plan, commit, reopen, diff, audit, hash, trace, and artifact evidence on a public fixture workbook. It needs a stable schema and honest "local evidence, not signed provenance" boundaries.
3. Feature fingerprint open planning can make first-window and agent-read workflows smarter. The probe classified formulas, pivots, slicers, macros, ActiveX, and simple value files without full model hydration.
4. Columnar sidecars should graduate to the performance loop. A 400k-cell numeric probe showed 10.87x faster repeated scans, or 5.59x including sidecar build, without replacing workbook semantics.
5. Oracle routing by mismatch class should graduate to the correctness loop. Cached values, accepted mismatches, Excel full rebuild, LibreOffice hard recalc, HyperFormula compatibility, and manual triage need explicit route fields in formula corpus output.
6. Budgeted agent view is product-ready as a contract. A projected summary cut one real sheet from about 5,844 approximate tokens to 792 while preserving shape, headers, column kinds, samples, and formula patterns.
7. Formula spans should become a language-service foundation. Tokens already have positions and parser coverage is strong, but AST spans and symbol roles are needed before safe rename.
8. MVCC viewport patch streams are a strong UI systems candidate. The current latest-token model is correct, but a retained delta ledger could patch from stale-but-retained tokens and expire compacted history explicitly.
9. Journal inverse properties should graduate as a correctness harness. The research probe found and folded in a real empty-cell style undo bug; the generated harness should become a shrinkable test suite.
10. Formula-logictest would make correctness claims easier to audit. The existing HyperFormula comparator already separates matches, known divergences, and skips; it needs a durable completed artifact.

## Dead Ends Or Hold-Backs

- A universal columnar workbook rewrite is still not justified. Columnar sidecars should target table/range scans and analytics, not replace `SparseGrid`.
- Token-only formula actions are not safe for rename. They are fine for hover/highlight, but code actions require AST spans plus binding resolution.
- Current write plan `origin` values are not enough for release proof language. They need action reasons and post-write evidence before being marketed as proof.
- Full raw `agentView` JSON is too large for strict agent budgets. Keep it, but add explicit budget modes instead of relying on callers to truncate.
- No single formula oracle is enough. Excel, LibreOffice, HyperFormula, cached values, and static goldens each need routed responsibility and explicit limitations.
- A release proof bundle must not imply SLSA/GitHub attestation unless signing and verifier roots actually exist.

## External Ideas To Keep

- HyperFormula's dependency graph documentation emphasizes range-node reuse for large repeated ranges; this supports Ascend's dependency graph and formula SOTA work.
- sqllogictest's completion/validation split is a direct fit for formula conformance artifacts.
- Apache Arrow and DuckDB point toward sidecar scan formats and replacement-table concepts, but only as secondary views over workbook truth.
- fast-check model-based testing is the right shape for journal laws because commands can be generated and shrunk.
- Univer and Notion show that agent-native spreadsheet/workspace tools are moving toward MCP, hosted tools, session identity, logs, CLI, and predictable execution surfaces.

## Fold-In Order

1. Correctness: `PackageDeltaPlan` action taxonomy and proof artifacts.
2. Product/DX and correctness: release proof bundle schema over existing plan/commit/post-write evidence.
3. Performance: feature fingerprint open planner with bounded package scans and timings.
4. Performance: private columnar range sidecar with generation invalidation and real-workbook benchmarks.
5. Correctness: formula oracle routing fields and counters in corpus JSON.
6. Product/DX: budgeted `agentView` contract across CLI/API/MCP.
7. Product/DX and performance: generation-aware viewport patch ledger.
8. Product/DX and correctness: AST spans plus symbol roles for formula diagnostics and rename previews.
9. Correctness: model-based journal inverse law harness.
10. Correctness: formula-logictest converter beside current JSON fixtures.

## Future Goal Prompts

```text
/goal Promote PackageDeltaPlan from research to a correctness loop. Do not change writer behavior first. Add an internal action taxonomy over existing write plans, prove it on scalar setCells, calc-chain drop, signature invalidation, and preserved docProps passthrough, then expose compact proof summaries in CLI/API/MCP tests.
```

```text
/goal Build a budgeted agent-view product/DX slice. Add a stable summary profile with byte/token budget controls across SDK, CLI, API, and MCP. Preserve deterministic truncation metadata and tests on a real formula/table workbook.
```

```text
/goal Build a formula-logictest correctness harness beside existing JSON fixtures. Start with a converter from current formula conformance files, emit completed artifacts with oracle metadata, and validate Ascend plus HyperFormula with explicit skip/known-divergence records.
```

```text
/goal Design parser-native formula spans. Add source ranges and symbol-role metadata without breaking current FormulaNode consumers, then prove hover/diagnostic ranges and rename-preview safety on structured refs, LET names, sheet spans, and external workbook refs.
```

```text
/goal Build a release proof bundle product/DX slice. Define a stable JSON schema over existing agent plan and commit results: subject, input hashes, operations, plan digest, trace/artifact digests, write policy, post-write reopen, diff, audits, and explicit claim boundaries. Add compact CLI/API output and full fixture-backed tests. Do not claim SLSA or signed provenance unless real attestations exist.
```

```text
/goal Promote formula oracle routing to the correctness loop. Add route fields and counters to formula corpus mismatch JSON without changing thresholds: accepted mismatch, Excel full rebuild, LibreOffice hard recalc, HyperFormula compatibility, static golden, and manual triage. Add tests for volatile, numeric drift, stale cache, unsupported function, external reference, dynamic array, structured reference, date-system, semantic, and oracle-error cases.
```

```text
/goal Prototype a private range scan sidecar in the performance loop. Keep SparseGrid authoritative. Add generation-keyed numeric/date column sidecars for stable rectangular ranges, validity bitmaps, checksums, invalidation reasons, and benchmark coverage on real workbook tables and agent-view scans.
```
