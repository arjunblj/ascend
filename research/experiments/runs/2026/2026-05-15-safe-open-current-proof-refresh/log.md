# Safe Open Current Proof Refresh

Date: 2026-05-15

## Question

Can the current top-ranked "safe unknown workbook opening" claim still be proven from existing SDK/CLI/API/MCP open-plan surfaces without adding a new opener or changing production behavior?

## Hypothesis

Yes. The tracked safe-open proof harness should still show package-feature routing before hydration, explicit review branches for active/security/unknown packages, malformed rejection, and materially lower local latency than full workbook hydration on public fixtures.

## External Sources Checked

- Microsoft Protected View frames unsafe file opening as a read-only/trust workflow, not a package-feature router: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions define package parts and relationships, which is the level Ascend inspects before workbook hydration: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft Open XML SDK documentation frames Open XML documents as inspectable Office document packages: https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk

## Why This Matters To Ascend

This remains the highest-ranked release claim because it is product-shaped, agent-native, and differentiates Ascend from spreadsheet libraries that start by hydrating workbook content. The proof must stay honest: Ascend routes open modes and review branches from package features; it does not scan malware or decide file trust.

## Probe/Implementation

No production code changed. Reran the existing tracked harness and surface tests:

```bash
bun run fixtures/benchmarks/safe-open-proof.ts --repeat 5 --warmup 1
bun test fixtures/benchmarks/safe-open-proof.test.ts packages/sdk/src/open-plan.test.ts
bun test apps/cli/src/cli.test.ts -t "open-plan"
bun test apps/api/api.test.ts -t "open-plan"
bun test apps/mcp/src/index.test.ts -t "open_plan|open-plan"
```

Updated `research/experiments/syntheses/2026-05-safe-open-proof-bundle.md` with the current proof table and validation status.

## Results

Latest proof run: 2026-05-15T03:55:56.646Z.

| Case | Mode | Review before hydration | Risk families | Median open-plan ms | Median full-open ms | Full/open-plan ratio | Boundary |
| --- | --- | --- | --- | ---: | ---: | ---: | --- |
| clean | formula | false | none | 0.200 | 2.297 | 11.46x | ok |
| formula-heavy | formula | false | none | 0.191 | 6.452 | 33.75x | ok |
| macro | metadata-only | true | preservedMacro | 0.166 | 1.460 | 8.80x | ok |
| pivot | formula | false | none | 0.145 | 2.544 | 17.56x | ok |
| ActiveX | metadata-only | true | preservedActiveX | 0.086 | 1.612 | 18.84x | ok |
| chart | formula | false | none | 0.077 | 1.248 | 16.27x | ok |
| signed | metadata-only | true | preservedSignature | 0.054 | 0.088 | 1.63x | synthetic package proof |
| unknown part | metadata-only | true | preservedOther | 0.037 | 0.076 | 2.06x | synthetic package proof |
| malformed | n/a | n/a | none | n/a | n/a | n/a | open-plan rejected: Missing end of central directory record |

Surface validation passed for SDK, CLI, API, and MCP open-plan behavior.

## Confidence

High for guarded release wording over the current fixtures and surfaces. Medium for publication because signed and unknown-part cases are still durable code-generated packages rather than public binary fixtures.

## Fold-In Decision

Promote to product/performance proof packaging only. Do not add another open-plan surface. The top implementation handoff remains publication/report packaging, plus optional replacement of synthetic signed and unknown-part cases with public binaries if the product loop requires that bar.

## Next Question

Can the auditable package-part mutation proof receive the same current refresh and stay aligned with the release proof index without adding mutation surfaces?
