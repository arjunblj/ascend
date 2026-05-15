# 2026-05 Release Decision Artifact

Updated: 2026-05-15

Decision state: two proof-backed release claims only. `headlineClaimsAllowed=false` and `implementationSurfacePromotionAllowed=false` remain the release gate until the listed owners close or explicitly reject their blockers.

## Claim 1: Safe Unknown Workbook Opening

Allowed wording today:

Ascend can inspect XLSX/XLSM package features before workbook hydration and produce a review/load recommendation for unknown or risky workbooks. The current proof covers public clean, formula-heavy, macro, pivot, ActiveX, and chart workbooks plus disclosed generated signed, unknown-part, and malformed packages.

Exact proof behind the claim:

| Evidence | Command or file | What it proves |
| --- | --- | --- |
| Safe-open proof harness | `bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json` | Case-level package-feature routing, recommended mode, review-before-hydration flag, risk families, malformed rejection, fixture provenance, and claim boundary text. |
| Safe-open proof tests | `bun test fixtures/benchmarks/safe-open-proof.test.ts` | Stable coverage over public, generated, and malformed cases without weakening publication blockers. |
| Public fixture scan | `bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json` | Current tracked corpus scan: 223 fixtures scanned, 0 signed/unknown replacements found, generated edge cases still owner-gated. |
| Release owner handoff | `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json` | `safe-open-proof` stays `blocked-by-publication-policy`; owner gates are explicit and machine-readable. |
| Latency profile gate | `fixtures/benchmarks/release-proof-index.ts` `safeOpenLatencyValidationEvidence` | Release latency requires repeat 10, warmup 3, public cases, timing environment, metrics, and CV guard; current default handoff is not timed release evidence. |

Claims we must not make:

- Malware scanning, sandboxing, file trust, active-content safety, Protected View equivalence, or recovery of arbitrary malformed packages.
- Signature validation, signer identity, signed provenance, SLSA, in-toto, Sigstore, GitHub artifact attestation, or tamper-evident storage.
- Release latency, SLA, hardware-normalized performance, or QSS performance comparison. The batched latency probe failed the CV guard and killed the simple batching fold-in.
- Public signed/unknown fixture coverage until product/release accepts generated topology or vendors approved public binaries.

A+ blockers and owner action:

| Owner | Blocking gate | Required next action |
| --- | --- | --- |
| Product | `public-edge-fixtures` | Accept disclosed generated signed/unknown topology for narrow proof wording, or replace with approved public binary fixtures. |
| Performance | `release-latency-run` | Run `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json` in a tracked-clean release environment over standardized public inputs; approve only non-threshold wording if CV guard and environment evidence are acceptable. |
| Release | `publication-boundary` | Approve copy that says pre-hydration package-feature routing only and excludes malware, sandbox, trust, active-content, signature, and recovery claims. |
| Release | `compact-report-publication-policy` | Define artifact storage, privacy filtering, canonical JSON subject, and offline verification expectations before publishing compact report digests. |

Decision: safe to present as a proof-backed internal/release-candidate claim. Not safe for headline public copy until owner gates close.

## Claim 2: Auditable Package-Part Mutation

Allowed wording today:

Ascend can produce local package-part mutation evidence that classifies workbook package changes as `passthrough`, `regenerate`, `add`, `drop`, or `error`, with source graph coverage, journal package issue refs, post-write audit status, and compact proof summaries. The current proof includes public docProps, calc-chain, macro, and chart fixtures, generated workbook add/regenerate cases, and disclosed generated signature/unknown edge packages.

Exact proof behind the claim:

| Evidence | Command or file | What it proves |
| --- | --- | --- |
| Package-action proof harness | `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json` | All five action classes, source graph evidence, package journal issue refs, post-write audit status, proof JSON sizing, and compact/full report boundaries. |
| Package-action proof tests | `bun test fixtures/benchmarks/package-action-proof.test.ts` | Stable proof shape, compact report boundary, all action classes, generated/public source counts, and five streaming proof cases. |
| Package fixture scan | `bun run fixtures/benchmarks/package-action-fixture-scan.ts --json` | Public replacements exist for several normal feature families, but signature and synthetic unknown path families remain 0 in the tracked corpus. |
| Release owner handoff | `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json` | `package-action-proof` stays `blocked-by-publication-policy`; owner gates, forbidden claims, and next actions are explicit. |
| Streaming matrix evidence | `fixtures/benchmarks/package-action-proof.ts` and `fixtures/benchmarks/release-proof-index.ts` | Five representative streaming proof cases now cover docProps, add-sheet, calc-chain, macro, and chart package accounting; generated edge/error streaming remains blocked. |

Claims we must not make:

- Full streaming parity across all package-action scenarios.
- Generated edge/error streaming parity: the unknown-part streaming probe did not emit the standard writer's fail-closed `error` action.
- Signature preservation, verification, re-signing, trust, signed provenance, SLSA, in-toto, Sigstore, or third-party attestation.
- Chart XML byte-passthrough, Excel recalculation equivalence, Excel-fresh cached formulas, macro/ActiveX safety, arbitrary unknown-part preservation, or semantic understanding of every unsupported feature.
- Public edge-fixture coverage while `signature-invalidation-drop` and `unknown-part-error` remain generated or owner-unapproved.

A+ blockers and owner action:

| Owner | Blocking gate | Required next action |
| --- | --- | --- |
| Product | `edge-fixture-policy` | Accept disclosed generated signature/unknown packages for guarded local proof wording, or replace them with approved public binary fixtures. |
| Correctness | `unsupported-feature-boundary` | Approve allowed/forbidden wording for signatures, calc chains, chart/drawing sidecars, macros/ActiveX, unknown parts, and streaming scope. |
| Performance | `streaming-matrix-boundary` | Accept the current five-case representative streaming matrix for narrow wording, or fund generated edge/error streaming semantics before any broader claim. |
| Release | `provenance-boundary` | Approve local-proof wording below signed provenance, SLSA, in-toto, Sigstore, GitHub artifact attestation, and tamper-evident storage. |
| Release | `compact-report-publication-policy` | Define artifact storage, privacy filtering, canonical JSON subject, and offline verification expectations before publishing compact report digests. |

Decision: safe to present as local auditable package-part mutation evidence. Not safe for public headline copy, full streaming parity, or provenance/trust wording.

## Implementation Priority

Do not add SDK, CLI, API, MCP, formula rename, token-view, viewport, columnar, oracle, or observability surfaces for this release block.

Only these actions change A+ status:

1. Product/release fixture decisions for generated signed/unknown edge cases.
2. Performance owner decision on safe-open latency evidence and package-action streaming wording.
3. Correctness owner approval of package-action unsupported-feature boundaries.
4. Release owner approval of publication boundaries and compact report policy.

Everything else is archived or proof-packaging-only until it changes one of those gates.
