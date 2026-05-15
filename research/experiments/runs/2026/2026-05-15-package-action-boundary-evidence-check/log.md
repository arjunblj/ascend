# Package Action Boundary Evidence Check

## Question

Can the package-action unsupported-feature boundary matrix be proved from current harness output, instead of living only as wording guidance?

## Hypothesis

Yes. The release proof index already runs both package-action and safe-open proofs. A derived evidence field can verify each boundary row while still keeping correctness owner approval required.

## External sources checked

- Microsoft Protected View documentation: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- OOXML calculation chain: https://ooxml.info/docs/12/12.3/12.3.1/
- Microsoft macro security settings: https://support.microsoft.com/en-gb/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373
- Microsoft ActiveX settings: https://support.microsoft.com/en-us/office/enable-or-disable-activex-settings-in-office-files-f1303e08-a3f8-41c5-a17e-b0b8898743ed
- SheetJS VBA blobs documentation: https://docs.sheetjs.com/docs/csf/features/vba/
- GitHub artifact attestations: https://docs.github.com/actions/concepts/security/artifact-attestations

## Why this matters to Ascend

Auditable package-part mutation is credible only if unsupported-feature language is backed by concrete cases. The valuable claim is not "Ascend understands every workbook feature"; it is "Ascend can account for package actions and refuse or route unsupported risk explicitly."

## Probe/implementation

Added `correctnessBoundaryEvidence` to `fixtures/benchmarks/release-proof-index.ts` and the compact owner handoff JSON. It derives six checks from current proof results:

- digital signatures: `signature-invalidation-drop` plus safe-open `signed`
- calc chain: `calc-chain-drop`
- chart/drawing sidecars: `chart-sidecar-accounting`
- macros/ActiveX: `macro-passthrough`, safe-open `macro`, and safe-open `activex`
- unknown parts: `unknown-part-error` plus safe-open `unknown-part`
- streaming scope: `docprops-passthrough` representative streaming proof

Each row copies the allowed/forbidden wording from `correctnessPolicy`, reports `evidencePresent`, and keeps `ownerApprovalRequired=true`.

## Results

Probe command:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json | jq '{headlineClaimsAllowed, implementationSurfacePromotionAllowed, missingRequirementCount, correctnessBoundaryEvidence}'
```

Observed:

- `headlineClaimsAllowed=false`
- `implementationSurfacePromotionAllowed=false`
- `missingRequirementCount=9`
- `correctnessBoundaryEvidence.status=evidence-present-owner-approval-required`
- `correctnessBoundaryEvidence.allCurrentEvidencePresent=true`
- `correctnessBoundaryEvidence.ownerApprovalRequired=true`
- all six feature checks report `evidencePresent=true`

Targeted test:

- `bun test fixtures/benchmarks/release-proof-index.test.ts`

## Confidence

High that the current boundary matrix is now machine-checkable. Medium for release wording, because correctness still needs to approve the boundary and product/release/performance gates remain open.

## Fold-in decision

Folded into release proof indexing as proof evidence only. Do not mark `unsupported-feature-boundary` satisfied, and do not promote any new mutation surface. This makes the correctness owner decision concrete.

## Next question

Can product resolve the generated edge fixture policy for the top two claims, or should correctness first approve the unsupported-feature boundary now that evidence is machine-checkable?
