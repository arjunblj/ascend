# Safe Open Gate Resolution Proof

## Question

Can the safe-open public-edge-fixture and release-latency gates be resolved without changing existing SDK, CLI, API, or MCP open surfaces?

## Hypothesis

Partially. The public-edge-fixture gate probably cannot be resolved by finding a checked-in public signed/unknown-part workbook replacement today, but it can be resolved by explicit product approval of disclosed generated structural fixtures. The release-latency gate should not be resolved from a one-off local timing run; it needs performance-owner threshold wording.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft OPC fundamentals and digital signatures: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft `System.IO.Packaging` namespace and package digital signatures: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging?view=windowsdesktop-10.0
- Microsoft Office macro blocking for untrusted files: https://learn.microsoft.com/en-us/deployoffice/security/internet-macros-blocked
- OOXML unknown relationships reference: https://ooxml.info/docs/9/9.1/9.1.7/
- DevExpress Office File API signing docs: https://docs.devexpress.com/OfficeFileAPI/405733/spreadsheet-document-api/document-security/sign-excel-files
- GemBox XLSX digital signature options docs: https://www.gemboxsoftware.com/spreadsheet/docs/GemBox.Spreadsheet.XlsxDigitalSignatureSaveOptions.html
- USENIX paper on OOXML signature security: https://www.usenix.org/system/files/usenixsecurity23-rohlmann.pdf

## Why this matters to Ascend

Safe unknown workbook opening is the top product/performance claim. If the fixture and latency gates stay vague, release copy can accidentally imply malware scanning, trust, or performance guarantees. The claim needs a crisp proof boundary: what is proven now, what is generated structural evidence, and what needs product or performance owner approval.

## Probe/implementation

- Inspected `fixtures/benchmarks/release-proof-index.ts`, `fixtures/benchmarks/safe-open-fixture-scan.ts`, the release claim board, and the ranked portfolio.
- Confirmed the tracked worktree was clean before the final safe-open probe with `git status --porcelain=v1 -uno`.
- Ran `bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json`.
- Ran `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json`.
- Updated `research/experiments/syntheses/2026-05-release-claim-board.md` with the gate resolution attempt.
- Finished and committed in-flight correctness fixes before returning to this synthesis:
  - `7dbaa6f2 fix(verify): catch detached shared formula members`
  - `d73c9177 fix(engine): retarget copied sheet formula text`

## Results

Fixture scan:

| Metric | Value |
| --- | ---: |
| Checked-in XLSX/XLSM fixtures scanned | 351 |
| Rejected during scan | 2 |
| Signed/unknown-part replacement candidates | 0 |
| Replacement status | `no-public-binary-replacement-found` |

Rejected fixtures were `fixtures/xlsx/calamine/pass_protected.xlsx` and `fixtures/xlsx/poi/protected_passtika.xlsx`.

Timed clean local safe-open proof:

| Case | Kind | Mode | Review | Open-plan median ms | Full-open median ms | Ratio |
| --- | --- | --- | --- | ---: | ---: | ---: |
| clean | public fixture | formula | false | 0.238 | 2.425 | 10.21 |
| formula-heavy | public fixture | formula | false | 0.244 | 9.181 | 37.66 |
| macro | public fixture | metadata-only | true | 0.099 | 1.595 | 16.15 |
| pivot | public fixture | formula | false | 0.168 | 3.162 | 18.77 |
| activex | public fixture | metadata-only | true | 0.113 | 1.998 | 17.70 |
| chart | public fixture | formula | false | 0.088 | 1.620 | 18.47 |
| signed | synthetic | metadata-only | true | 0.048 | 0.148 | 3.11 |
| unknown-part | synthetic | metadata-only | true | 0.046 | 0.088 | 1.91 |

Malformed synthetic bytes were rejected with `Missing end of central directory record`.

## Confidence

High that checked-in public fixtures do not currently contain signed/unknown replacements. Medium that no suitable public binary exists elsewhere; web search found signing APIs and documentation more readily than durable public sample binaries. High that the timed local proof is useful diagnostic evidence but insufficient for release latency wording.

## Fold-in decision

Promote to topic synthesis and owner-loop handoff only. Do not change open surfaces. Keep `public-edge-fixtures` and `release-latency-run` missing in the release proof index until:

- Product explicitly accepts disclosed generated signed/unknown structural fixtures or supplies public binary replacements.
- Performance approves the release latency environment, repeat count, input set, and threshold wording.

## Next question

Can the package-action proof owner resolve the `unsupported-feature-boundary` gate by producing a concise correctness boundary matrix for signatures, chart byte passthrough, Excel recalculation equivalence, unknown parts, and streaming coverage?
