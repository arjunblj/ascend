# Release Claim Evidence Cut

## Question

Can the current release-claim board distinguish claims that are already backed by repeatable proof harnesses from claims that still need publication-strength evidence, without promoting another product surface?

## Hypothesis

Yes. The top two claims now have tracked harness evidence. The right move is to update the board so "proof still missing" means remaining release/publication proof, not implementation or harness work that already exists.

## External sources checked

- Open Packaging Conventions package parts, relationships, and signatures: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft Excel digital signatures and invalidation after saving a copy: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing
- SheetJS CE write options and writer preservation boundaries: https://docs.sheetjs.com/docs/api/write-options/
- openpyxl preservation boundaries and unsupported object warnings: https://openpyxl.readthedocs.io/en/stable/tutorial.html

## Why this matters to Ascend

The release board is supposed to constrain implementation, not encourage another surface. If it keeps saying the top claims lack proof after tracked harnesses exist, future loops may add redundant APIs instead of turning the evidence into release-ready artifacts with honest wording.

## Probe/implementation

- Ran `git status --short --branch` before edits; unrelated writer files are dirty and were not touched.
- Inspected current safe-open and package-action proof harnesses:
  - `fixtures/benchmarks/safe-open-proof.ts`
  - `fixtures/benchmarks/package-action-proof.ts`
  - associated tests and prior experiment logs.
- Reran current proof commands:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1
bun run fixtures/benchmarks/package-action-proof.ts
```

- Updated the release claim board and claim ladder language to separate:
  - repeatable local proof harnesses that exist today;
  - missing publication evidence, such as release-environment reruns, durable public binary fixtures for synthetic edge cases, and a report artifact generated from existing surfaces.

## Results

Safe unknown workbook opening:

- 9 proof cases ran: clean, formula-heavy, macro, pivot, ActiveX, chart, signed, unknown-part, malformed.
- Public active-content fixtures route to `metadata-only` with `reviewBeforeHydration: true`.
- Malformed bytes are rejected explicitly.
- Local median full-open/open-plan ratios ranged from `2.01x` to `45.37x`; these are proof-run numbers, not release thresholds.

Auditable package-part mutation:

- 8 proof cases ran.
- Combined commit evidence covered every action kind: `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1`.
- Every case included source graph evidence and output byte digests.
- The unknown-part case produced the expected post-write audit issue.
- The chart fixture remains a boundary: chart XML is regenerated while drawing sidecars pass through, so the claim must say accounting, not chart byte passthrough.

Board update:

- Safe-open is now marked proof-backed for conservative wording, with publication proof still missing.
- Package-action is now marked proof-backed for local action taxonomy, with publication fixtures/reporting still missing.
- Formula intelligence remains no-rename and rejection-first.
- Columnar sidecars and safe formula rename remain do-not-promote.

## Confidence

High that the board now reflects current evidence. Medium that the two top claims are publication-ready without another pass: both still depend partly on synthetic edge packages, and timing numbers should be rerun in the release environment.

## Fold-in decision

Promote to topic synthesis only. Do not add product surfaces. Hand off the next work as release proof packaging over existing harnesses and surfaces.

## Next question

Can Ascend generate a single release proof index that references safe-open and package-action harness outputs by digest, without embedding bulky artifacts or implying signed provenance?
