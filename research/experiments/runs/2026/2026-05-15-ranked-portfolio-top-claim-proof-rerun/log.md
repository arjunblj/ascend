# Ranked Portfolio Top Claim Proof Rerun

Date: 2026-05-15

## Question

Can the ranked research portfolio stop broad exploration and prove the top one or two product-shaped claims from existing surfaces?

## Hypothesis

Yes. The highest-leverage unknowns are not new feature ideas; they are whether the top two release claims still have current, repeatable proof: safe unknown workbook opening and auditable package-part mutation.

## External sources checked

- Microsoft Protected View frames untrusted Office documents as read-only/protected review rather than trust by default: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions frame XLSX-style documents as parts plus relationships, which is the right proof boundary for pre-hydration inspection and per-part mutation evidence: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- openpyxl documents preservation boundaries for unsupported workbook objects, including `keep_vba` and unsupported shapes: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- SheetJS write options document explicit writer support boundaries for spreadsheet features: https://docs.sheetjs.com/docs/api/write-options/
- DuckDB documents direct XLSX range/table import, useful contrast for keeping columnar sidecars out of the top release proof index: https://duckdb.org/docs/stable/guides/file_formats/excel_import

## Why this matters to Ascend

Ascend's North Star depends on claims humans and agents can trust. The top two ranked claims are product-shaped and differentiating, but only if they remain proof-backed without scope creep: package-feature safe opening before hydration, and auditable package-part mutation with explicit outcomes.

## Probe/implementation

No new product surface was added.

- Reran `fixtures/benchmarks/safe-open-proof.ts` with timings.
- Reran `fixtures/benchmarks/package-action-proof.ts` with timings.
- Reran targeted harness tests for both artifacts.
- Reran the release proof index over only the top two artifacts.
- Updated the ranked portfolio and the two proof reports with current evidence.
- Finished and committed the in-flight data-validation journal order exactness fix separately before returning to synthesis.

## Results

Safe unknown workbook opening:

- 9 proof cases.
- 8 OK cases and 1 malformed rejection.
- 4 cases route to `reviewBeforeHydration`.
- Public fixture open-plan speedup range in this local rerun: 11.70x to 39.77x versus full hydration.
- Synthetic signed and unknown-part packages route to metadata-only review.

Auditable package-part mutation:

- 8 proof cases.
- Combined commit actions: `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1`.
- Every case includes source graph evidence.
- Every case has a package-preservation journal issue.
- The unknown-part case has the expected package-action proof issue and remains review-required.

Release proof index:

- `safe-open-proof` stable shape digest: `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`.
- `package-action-proof` stable shape digest: `b9758496346c97920c80ba08b6632315708a6d6cc770927695337e729554dbb0`.
- The index is explicitly unsigned and not an attestation.

Validation passed:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1
bun run fixtures/benchmarks/package-action-proof.ts
bun test fixtures/benchmarks/safe-open-proof.test.ts packages/sdk/src/open-plan.test.ts
bun test fixtures/benchmarks/package-action-proof.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings
```

## Confidence

High that the top two handoffs are the right implementation loops for the next block. Medium that synthetic signed/unknown cases are publication-ready; the proof is repeatable, but product may still want durable binary fixtures before external release copy.

## Fold-in decision

Promote to topic synthesis and handoff only. Do not add new surfaces. The next implementation loops should package the proof artifacts for release publication using existing SDK/CLI/API/MCP evidence.

## Next question

Can product/performance publish the safe-open proof bundle and correctness/product publish the package-action proof bundle without adding another narrow command or broadening claim language?
