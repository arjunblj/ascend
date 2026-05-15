# Release Correctness Policy Checklist

## Question

Can the unsupported-feature boundary for auditable package-part mutation become a machine-readable correctness owner checklist without promoting a new mutation surface or marking the release gate satisfied?

## Hypothesis

Yes. The existing package-action proof and release claim board already define the allowed/forbidden wording. Folding that matrix into `release-proof-index` should make the owner decision explicit while keeping `unsupported-feature-boundary` missing until correctness approves it.

## External sources checked

- OOXML calculation chain: https://ooxml.info/docs/12/12.3/12.3.1/
- Microsoft macro security settings: https://support.microsoft.com/en-gb/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373
- Microsoft ActiveX settings: https://support.microsoft.com/en-us/office/enable-or-disable-activex-settings-in-office-files-f1303e08-a3f8-41c5-a17e-b0b8898743ed
- SheetJS VBA blobs: https://docs.sheetjs.com/docs/csf/features/vba/
- OOXML digital signatures: https://c-rex.net/samples/ooxml/e1/Part2/OOXML_P2_Open_Packaging_Conventions_Digital_topic_ID0EHROM.html

## Why this matters to Ascend

Ascend should be able to claim auditable package-part mutation only when the wording is precise: per-part accounting and fail-closed behavior are credible; signature verification, macro safety, chart semantics, Excel-fresh cached formulas, arbitrary unknown-part understanding, and full streaming parity are not proven.

## Probe/implementation

Added `correctnessPolicy` to `fixtures/benchmarks/release-proof-index.ts` and the compact owner handoff JSON. The policy includes:

- one pending `package-action-proof/unsupported-feature-boundary` approval item for the correctness loop
- a six-row unsupported-feature matrix for digital signatures, calc chain, chart/drawing sidecars, macros/ActiveX, unknown parts, and streaming scope
- source references and an explicit boundary that this is wording approval, not semantic workbook support

The existing `readyWhen` gate remains missing and `headlineClaimsAllowed` remains false.

## Results

Targeted validation:

- `bun test fixtures/benchmarks/release-proof-index.test.ts`

The test now asserts that `correctnessPolicy` is present in the release proof index, compact handoff JSON, and Markdown output, and that it does not satisfy the release gate.

## Confidence

Medium-high. This is a small evidence-routing change over an existing proof artifact. It strengthens owner handoff clarity without changing workbook behavior, mutation behavior, or release readiness.

## Fold-in decision

Folded into release proof indexing as owner-decision metadata. Do not promote a production mutation surface from this change. Correctness must still approve or reject the wording before `unsupported-feature-boundary` can move from missing to satisfied.

## Next question

Which top claim should receive the next proof-producing loop: resolving the safe-open generated fixture approval policy, or resolving the package-action unsupported-feature wording with a correctness owner decision?
