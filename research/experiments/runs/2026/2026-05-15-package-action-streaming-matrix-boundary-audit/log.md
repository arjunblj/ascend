# Package Action Streaming Matrix Boundary Audit

## Question

Can the package-action `streaming-matrix-boundary` gate become approval-ready without expanding the writer surface, or does release wording need a broader streaming proof matrix?

## Hypothesis

One representative streaming proof is enough to document a narrow dirty-sheet passthrough claim, but not enough to support release wording that implies full streaming parity across package-action classes or feature families.

## External sources checked

- openpyxl optimized read/write modes: https://openpyxl.readthedocs.io/en/stable/optimized.html
- Excelize StreamWriter docs: https://xuri.me/excelize/en/stream.html
- JSZip `generateAsync` streaming docs: https://stuk.github.io/jszip/documentation/api_jszip/generate_async.html
- JSZip file data docs: https://stuk.github.io/jszip/documentation/api_jszip/file_data.html

## Why this matters to Ascend

Auditable package-part mutation is a top release claim. If the release copy mentions streaming without a matrix boundary, readers may infer that every package action and unsupported-feature case has streaming parity. Current evidence does not support that.

## Probe/implementation

- Inspected `fixtures/benchmarks/package-action-proof.ts`, `packages/io-xlsx/src/writer/index.ts`, `packages/io-xlsx/src/writer/plan.ts`, and streaming writer tests.
- Ran `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json`.
- Ran a local Bun audit to summarize streaming coverage by case and action class.
- Updated `research/experiments/syntheses/2026-05-release-claim-board.md` with a streaming matrix boundary.

## Results

Local streaming audit:

| Metric | Value |
| --- | --- |
| Total package-action cases | 8 |
| Cases with streaming proof | 1 |
| Streaming proof case | `docprops-passthrough` |
| Non-streaming cases | `regenerate-existing-sheet`, `add-sheet-part`, `calc-chain-drop`, `signature-invalidation-drop`, `macro-passthrough`, `chart-sidecar-accounting`, `unknown-part-error` |
| Package-action classes covered overall | `passthrough`, `regenerate`, `add`, `drop`, `error` |
| Package-action classes covered by streaming proof | `passthrough`, `regenerate` |
| Streaming regenerated parts | `xl/worksheets/sheet1.xml` |
| Streaming passthrough byte-equal count | 3 |

Boundary matrix:

| Boundary | Allowed claim | Forbidden claim |
| --- | --- | --- |
| Representative streaming passthrough | one streaming dirty-sheet write preserves passthrough parts while regenerating the dirty worksheet | streaming proof covers every action class |
| Add/drop/error actions | non-streaming proof covers add/drop/error | streaming add/drop/error parity is proven |
| Macro/chart public fixtures | public macro/chart package accounting is proven in the standard writer path | streaming macro/chart preservation is release-proven |
| ZIP streaming semantics | streaming wording is limited to the tested dirty-sheet passthrough case | streaming mode is semantically equivalent across all workbook/package features |

## Confidence

High that current proof supports only narrow streaming wording. Medium on whether one representative case is enough for release copy; that is a performance-owner decision.

## Fold-in decision

Promote to topic synthesis only. Keep `streaming-matrix-boundary` missing in the release proof index. Do not add writer surfaces. A future performance loop should either accept the narrow wording or expand the package-action proof matrix to streaming variants before any broader streaming claim.

## Next question

Can the compact-report publication gate be reduced to an owner policy decision, or is there a real missing proof around artifact storage, privacy filtering, and canonicalization?
