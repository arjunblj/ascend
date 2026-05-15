# Correctness Boundary Diagnostics

## Question

Can unsupported-feature boundary proof regressions become machine-readable owner blockers instead of only flipping a broad boolean?

## Hypothesis

Yes. The release proof index already computes feature-level evidence for signatures, calc chain, chart/drawing sidecars, macros/ActiveX, unknown parts, and streaming scope. It can expose the failed feature names and an escalation flag while preserving the existing owner-approval gate.

## External sources checked

- Microsoft OOXML calculation chain reference: https://ooxml.info/docs/12/12.3/12.3.1/
- Microsoft macro security settings: https://support.microsoft.com/en-gb/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373
- OOXML digital signatures in Open Packaging Conventions: https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Digital_topic_ID0EHROM.html

## Why this matters to Ascend

The auditable package-part mutation claim depends on honest unsupported-feature boundaries. A regression in signature, calc-chain, macro, chart, unknown-part, or streaming evidence should tell correctness owners exactly which claim boundary lost proof. A single `allCurrentEvidencePresent=false` is not enough for a release owner to route the work.

## Probe/implementation

- Inspected `correctnessBoundaryEvidence` in `fixtures/benchmarks/release-proof-index.ts`.
- Ran the owner handoff JSON probe:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json | jq '{headlineClaimsAllowed, implementationSurfacePromotionAllowed, missingRequirementCount, correctnessBoundaryEvidence}'
```

- Added `missingFeatureNames` and `ownerEscalationRequired` to `correctnessBoundaryEvidence`.
- Rendered both fields in release-proof Markdown.
- Added regression assertions in `fixtures/benchmarks/release-proof-index.test.ts`.

## Results

Current proof:

| Field | Value |
| --- | --- |
| All current evidence present | `true` |
| Missing feature names | `[]` |
| Owner escalation required | `false` |
| Owner approval required | `true` |

Validation:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
```

Result: 4 tests passed.

## Confidence

High for the diagnostic shape because it is derived from the existing feature checks. Medium for future regression triage because the feature-level checks still depend on current harness behavior and should remain backed by targeted proof commands.

## Fold-in decision

Promote to correctness and release loops as owner-handoff evidence only. This does not satisfy `unsupported-feature-boundary`; correctness still owns approval of allowed/forbidden wording.

## Next question

Can release readiness expose a compact "claim blocker board" grouped by claim and owner so top handoff loops do not need to infer priorities from multiple evidence objects?
