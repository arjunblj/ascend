# Safe Open Risk Family Scan Evidence

## Question

Can the safe-open public-edge-fixture blocker report stronger tracked-corpus evidence than "zero signed/unknown matches" without pretending the product owner has accepted generated fixtures?

## Hypothesis

Yes. If the fixture scan counts every safe-open risk family it sees, then the owner handoff can show that the tracked corpus does contain detectable package-risk features such as macro and ActiveX workbooks while still showing zero tracked signed or unknown-part replacements.

## External sources checked

- Microsoft digital signatures and code signing in workbooks: https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/digital-signatures-code-signing
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- GitHub repository license API: https://docs.github.com/en/rest/licenses/licenses
- OpenSSF Scorecard binary artifacts check: https://github.com/ossf/scorecard/blob/main/docs/checks.md

## Why this matters to Ascend

Safe unknown workbook opening is the rank-1 claim, but its product blocker is fixture policy: signed and unknown-part cases are generated structural packages today. A stronger scan should help the product owner decide whether to accept disclosed generated topology proof or require public binary replacements. It must not close the gate automatically.

## Probe/implementation

Folded a tiny proof-harness improvement into `fixtures/benchmarks/safe-open-fixture-scan.ts`:

- count every risk family detected while scanning tracked public XLSX/XLSM fixtures;
- include `riskFamilyCounts` in scan JSON and Markdown;
- propagate safe-open risk-family counts into `release-proof-index` fixture policy evidence;
- add tests that the tracked corpus detects macro/ActiveX-style risk families while still reporting zero signed/unknown replacements.

## Results

Focused validation passed:

```bash
bun test fixtures/benchmarks/safe-open-fixture-scan.test.ts fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/safe-open-fixture-scan.ts fixtures/benchmarks/safe-open-fixture-scan.test.ts fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bunx tsc --build
```

Probe output:

```json
{
  "scanned": 223,
  "rejected": 1,
  "riskFamilyCounts": {
    "preservedActiveX": 2,
    "preservedControl": 3,
    "preservedMacro": 2,
    "preservedVendorSecurity": 1
  },
  "signatureOrUnknownMatches": 0,
  "replacementStatus": "no-public-binary-replacement-found"
}
```

This strengthens the public-edge-fixtures owner handoff: the tracked corpus scan can detect known risk families, but it still finds no signed or unknown-part public binary replacement.

## Confidence

High that the scan now gives better owner-decision evidence for the tracked corpus. Medium that this materially changes the product decision, because external public fixture discovery, license review, and binary provenance remain owner-owned.

## Fold-in decision

Promote to the release-proof harness and product-owner handoff. Keep `public-edge-fixtures` missing. Do not claim that generated signed/unknown fixtures are accepted, public, or equivalent to real-world signed workbooks.

## Next question

Can the safe-open proof owner define a public-fixture acceptance checklist that separates topology-only generated packages from real-world signed workbook behavior without adding a new open surface?
