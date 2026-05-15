# Compact Report Publication Policy Proof

## Question

Can release approve compact proof reports as reproducibility pointers without publishing workbook bytes, per-part action rows, generated artifacts, or local digests as provenance?

## Hypothesis

Yes, but only as a policy decision. Current compact reports already omit sensitive/heavy proof payloads and keep release gates blocked. They can support owner review, not public artifact publication, until storage, privacy filtering, and canonicalization are approved.

## External sources checked

- SLSA provenance defines signed provenance expectations that compact local reports do not satisfy: https://slsa.dev/spec/v1.0-rc1/provenance
- Open Packaging Conventions define the package-part layer whose detailed rows are intentionally omitted from compact package-action reports: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft Protected View anchors the safe-open trust boundary: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653

## Why this matters to Ascend

Both top claims are blocked on `compact-report-publication-policy`. Research should provide evidence about current compact report shape, then stop. Release owners should decide storage, privacy filtering, canonicalization, and whether compact report commands are enough for reproducibility.

## Probe/implementation

Ran:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
```

Searched compact JSON output for workbook bytes, input/output bytes, SHA-256 digests, and per-part digest fields. No production code changed.

## Results

- Safe-open compact report boundary says it omits workbook bytes and input digests.
- Safe-open compact output did not contain `inputSha256`, `inputBytes`, `outputBytes`, or per-part digest fields.
- Package-action compact report boundary says it omits workbook bytes, per-part action rows, and generated artifacts.
- Package-action compact output did not contain input/output workbook byte fields or per-part digest fields.
- Owner handoff still reports `implementationSurfacePromotionAllowed=false` and keeps `compact-report-publication-policy` blocked for both top artifacts.

## Confidence

High for current compact report field shape. Medium for release policy because approval still depends on where compact reports would be stored and whether generated fixture labels are enough for public release artifacts.

## Fold-in decision

Promote to topic synthesis only. Recommended boundary: compact reports may be used as local owner-review pointers today, but do not publish compact report digests or claim attestation until release defines storage, privacy filtering, and canonicalization policy.

## Next question

Is there any remaining top-claim blocker that requires research implementation, or are the remaining blockers all owner approvals and validation runs?
