# Copied Shared Formula Binding Trust Moat

## Question

Does the current quality moat prove copied shared-formula bindings persist as retargeted metadata, not only as normalized formula text?

## Hypothesis

Yes. The existing copy-sheet trust moat validates reopened formulas, adjacent metadata rewrites, and direct shared-formula binding assertions.

## External sources checked

None. This was a local validation-only cycle over the current SDK test assertion.

## Why this matters to Ascend

Auditable mutation claims depend on saved workbooks reopening with trustworthy formula metadata. Formula text alone is weaker evidence than formula text plus retargeted shared-formula binding ranges.

## Probe/implementation

Validated the current SDK test assertion in `packages/sdk/src/agent-workflow.test.ts`. The test checks that copied shared formula cells reopen with:

- master binding `A4`, range `Copy!A4:A5`;
- member binding `A4`, range `Copy!A4:A5`;
- existing data-validation and conditional-format formulas retargeted to `Copy`.

## Results

Focused validation:

```bash
bun test packages/sdk/src/agent-workflow.test.ts -t "quality moat matrix"
```

Result: 1 pass, 0 fail, 33 expect calls.

## Confidence

High for the copied shared-formula trust moat assertion. This is regression evidence only; it does not promote formula rename or a new product surface.

## Fold-in decision

Promote to correctness-loop evidence. No production code change in this cycle.

## Next question

Return to claim stewardship: should package-action proof swap the calc-chain generated edge case to a public fixture, or keep that as a future owner-owned harness change?
