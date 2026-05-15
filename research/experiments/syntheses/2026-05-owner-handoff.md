# 2026-05 Owner Handoff

Date: 2026-05-15

## Decision

Hand off only two implementation loops:

1. **Safe unknown workbook opening** to product/performance.
2. **Auditable package-part mutation** to correctness/product, with release and performance boundary approvals.

Everything else remains do-not-promote for this block. Formula intelligence stays rejection-first; do not implement rename.

Promotion throttle: current `release-proof-index` still reports `headlineClaimsAllowed=false`, `implementationSurfacePromotionAllowed=false`, and 9 missing owner requirements. Until those owner gates move, research may run diagnostic probes and record dead ends, but it should not promote new product claims or new public surfaces beyond the two handoffs below.

Latency diagnostic note: `practical-latency-contracts` now reports input provenance separately from tracked-code cleanliness. The `public-tracked` preset can generate a tracked-harness edit workbook for edit/verify phases, and labels it as generated evidence. This helps performance owners avoid turning private-corpus latency numbers into release claims, but generated benchmark inputs remain below public real-workbook fixture evidence.

Fixture-search note: the targeted checked-in fixture scan still finds no public binary replacement for safe-open signed or unknown-part structural cases. Latest scan: 351 fixtures scanned, 2 protected fixtures rejected, 0 signature/unknown matches. Product must accept disclosed generated structural fixtures or provide public binaries.

Compact privacy note: field-level inventory of the two compact reports finds no workbook bytes or proof digest artifact fields. The compact reports still include command strings and public fixture paths, so they remain local proof summaries until release owners define publication storage, privacy filtering, canonicalization, and verification expectations.

Machine-readable handoff note: the release index now emits `readiness.implementationHandoffs` for the top two claims, and each handoff includes `proofRequired` with fixture, benchmark, existing-surface boundary, validation gate, competitor contrast, honest boundary, and kill criterion. The same JSON result also emits `deferredClaims` for lower-ranked directions that should not promote in this block.

Streaming proof note: the release index now emits `streamingMatrixEvidence` for the package-action performance gate. Current proof covers one streaming case, `docprops-passthrough`, and action kinds `passthrough` and `regenerate`; it explicitly leaves `add`, `drop`, `error`, public macro, and public chart cases unproven for streaming. The gate remains owner-approval required.

## Proof Snapshot

Current local proof gate:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

| Field | Value |
| --- | --- |
| Release artifacts | `safe-open-proof`, `package-action-proof` |
| Headline claims allowed | `false` |
| Missing requirements | 9 |
| Top product gates | `package-action-proof/edge-fixture-policy`, `safe-open-proof/public-edge-fixtures` |
| Top correctness gate | `package-action-proof/unsupported-feature-boundary` |
| Top performance gates | `safe-open-proof/release-latency-run`, `package-action-proof/streaming-matrix-boundary` |
| Top release gate | `package-action-proof/provenance-boundary` |
| Implementation surface promotion allowed | `false` |
| Deferred claims | 6 |

Latest compact proof refresh:

| Artifact | Stable shape | Headline allowed | Key coverage |
| --- | --- | --- | --- |
| `safe-open-proof` | `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178` | false | cases=9, rejected=1, reviewBeforeHydration=4, public=6, synthetic=2, malformed=1 |
| `package-action-proof` | `b4cf0755f2fd40ff577ab87b6600ee14a58bbd69608796c903759e7e8b4d29e8` | false | cases=8, passthrough=32, regenerate=40, add=3, drop=3, error=1, streamingProofCases=1 |

## Owner Board

| Owner loop | Claim | Proof required next | Exit criterion |
| --- | --- | --- | --- |
| Product/performance | Safe unknown workbook opening | Accept or replace generated signed/unknown fixtures; run release-environment open-plan latency on public inputs; approve wording that excludes malware scanning, trust, active-content safety, signed provenance, and release threshold claims. | Claim copy may say pre-hydration package-feature routing and review recommendation across existing surfaces. |
| Correctness/product | Auditable package-part mutation | Accept disclosed generated edge-package evidence or replace it; approve unsupported-feature matrix for signatures, calc chain, chart/drawing sidecars, macros/ActiveX, unknown parts, and streaming; preserve journal/package issue compatibility. | Claim copy may say local per-part `passthrough`/`regenerate`/`add`/`drop`/`error` evidence with journal-linked package issues. |
| Performance | Package-action streaming boundary | Accept one representative streaming proof as sufficient for narrow wording or expand the matrix before broader wording. | Release copy may mention one representative streaming dirty-sheet proof only; no full streaming parity. |
| Release | Provenance and compact report policy | Approve local-proof wording; define artifact storage, retention/privacy filtering, canonicalization, and verification expectations. | Compact commands remain pointers until this exists; no SLSA, in-toto, Sigstore, GitHub attestation, or signed provenance wording. |

## Safe-Open Acceptance Checkboxes

Current proof input:

| Field | Value |
| --- | --- |
| Proof cases | 9 |
| Public fixture cases | 6 |
| Generated edge-package cases | 2 |
| Malformed cases | 1 |
| Generated/disclosed cases | `signed`, `unknown-part`, `malformed` |
| Review before hydration | 4 |
| Malformed rejected | true |

Owner acceptance checklist:

- [ ] Product accepts disclosed generated `signed` and `unknown-part` structural fixtures, or replaces them with public binary fixtures.
- [ ] Product accepts generated malformed-package evidence as structural rejection proof, not real-world repair behavior.
- [ ] Performance runs tracked-clean release-environment latency over standardized public inputs and approves wording that does not read as a threshold unless thresholds are explicitly owned.
- [ ] Release approves boundary wording: not malware scanning, sandboxing, file trust, active-content safety, signed provenance, or malformed-package recovery.
- [ ] Release keeps compact report digests unpublished until storage, privacy filtering, canonicalization, and verification expectations exist.

## Package-Action Acceptance Checkboxes

Current proof input:

| Field | Value |
| --- | --- |
| Proof cases | 8 |
| Public fixture cases | 4 |
| Generated workbook cases | 2 |
| Generated edge-package cases | 2 |
| Generated/disclosed cases | `regenerate-existing-sheet`, `add-sheet-part`, `signature-invalidation-drop`, `unknown-part-error` |
| Action classes | `passthrough=32`, `regenerate=40`, `add=3`, `drop=3`, `error=1` |
| Source graph evidence everywhere | true |
| Journal package issues everywhere | true |
| Post-write audit failures | `unknown-part-error` |
| Representative streaming proof cases | 1 |

Owner acceptance checklist:

- [ ] Product accepts disclosed generated edge packages for docProps passthrough, calc-chain drop, signature invalidation drop, and unknown-part error, or replaces them with public binary fixtures.
- [ ] Correctness approves unsupported-feature boundaries: signatures are detected/reported but not verified or re-signed; calc-chain decisions are package accounting, not Excel recalc equivalence; chart/drawing sidecars are accounted separately, not semantic chart support; macros/ActiveX are package evidence, not safety; unknown parts fail closed.
- [ ] Correctness keeps journal/package issue compatibility as part of the claim: every case must include a package-preservation journal issue.
- [ ] Performance accepts one representative streaming dirty-sheet proof as sufficient for narrow wording, or expands streaming variants before any broader streaming claim.
- [ ] Release approves local-proof wording that excludes SLSA, in-toto, signed provenance, third-party attestation, and tamper-evident storage.
- [ ] Release keeps compact report digests unpublished until storage, privacy filtering, canonicalization, and verification expectations exist.

## Source Of Truth

`fixtures/benchmarks/release-proof-index.ts` remains the machine source of truth for release gates. This handoff is a human approval checklist. Do not add a second machine-checked Markdown acceptance table here: the current handoff intentionally has 11 owner checkboxes split from 9 canonical `readyWhen` gates, so treating the Markdown as canonical would duplicate and drift from the release index.

Owner loops should consume these JSON fields first:

| JSON field | Purpose |
| --- | --- |
| `readiness.implementationSurfacePromotionAllowed` | Fail-closed guard for new SDK/CLI/API/MCP surfaces. Currently `false`. |
| `readiness.implementationHandoffs` | Canonical top-two handoffs with owner loops, proof commands, blocker IDs, next-step kinds, and proof requirements. |
| `readiness.implementationHandoffs[].proofRequired` | Product-shaped proof ladder: fixture, benchmark, surface, validation gate, competitor contrast, honest boundary, and kill criterion. |
| `fixturePolicyEvidence` | Summarizes tracked safe-open and package-action fixture scans for product gate decisions while keeping public replacement gaps explicit. |
| `correctnessBoundaryEvidence` | Verifies the unsupported-feature boundary matrix against current package-action and safe-open proof cases while keeping owner approval required. |
| `streamingMatrixEvidence` | Verifies the representative streaming proof boundary: covered action kinds, missing action kinds, covered case names, non-streaming public cases, and owner approval requirement. |
| `compactReportPublicationEvidence` | Proves compact report commands and privacy/canonicalization blockers without indexing compact report digests. |
| `deferredClaims` | Machine-readable do-not-promote/proof-backed-hold list for non-top directions. |
| `excludedEvidence` | Evidence that exists but must not become release proof yet, currently practical latency contracts. |

Owner-loop shortcut:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json
```

This emits the routing fields above without embedding full proof artifacts.

## Do Not Promote

| Direction | Freeze reason |
| --- | --- |
| Formula language-service primitives | Listed in `deferredClaims` as `proof-backed-hold`; rejection-first proof exists, but rename remains unsafe without workbook-context symbol ownership and operation-owned edits. |
| Retained viewport patch history | Listed in `deferredClaims` as `proof-backed-hold`; product-proof backed, but not a top release handoff and not collaboration/CRDT evidence. |
| Token-bounded agent view | Listed in `deferredClaims` as `proof-backed-hold`; product-proof backed, but token counts are approximate and omitted evidence is by design. |
| Release proof bundle | Ingredients exist, but artifact storage/privacy/canonicalization and non-attestation policy are unresolved. |
| Formula conformance/oracle routing | Listed in `deferredClaims` as `do-not-promote-yet`; still a correctness research program; no broad Excel-compatible formula claim. |
| Property-style journal laws | Useful correctness evidence, but not a release claim until coverage and shrinkability are owner-approved. |
| Columnar scan sidecars | Listed in `deferredClaims` as `do-not-promote-yet`; performance research only; not a product surface. |
| Agent workflow observability | Listed in `deferredClaims` as `do-not-promote-yet`; supporting DX evidence, not a top release claim. |

Formula freeze status: latest proof command `bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json` sampled 1685 public formulas, found 2322 reference spans, 25 binding roles, 3 LET-local prepare-rename OK targets, and 1692 prepare-rename refusals. This remains a product/DX proof of formula-local assist and rejection classification only. Do not implement edit-producing rename; do not promote workbook-context defined-name, table, external, 3D, spill, sheet, cell, or range rename.

## Next-Loop Prompts

### Product/Performance

```text
/goal Become Ascend's safe unknown workbook opening proof owner. Do not add new open surfaces. Start from `fixtures/benchmarks/release-proof-index.ts` and `safe-open-proof`. Resolve or explicitly accept generated signed/unknown package fixtures, run release-environment open-plan latency on standardized public inputs, and approve wording that says pre-hydration package-feature routing and review recommendation only. Do not claim malware scanning, sandboxing, trust, active-content safety, signed provenance, or release latency thresholds. Commit only proof/report/index changes or narrow harness fixes.
```

### Correctness/Product

```text
/goal Become Ascend's auditable package-part mutation proof owner. Do not add mutation surfaces. Start from `fixtures/benchmarks/release-proof-index.ts` and `package-action-proof`. Resolve generated edge-package policy, approve the unsupported-feature matrix, preserve journal/package issue compatibility, and keep streaming/provenance boundaries explicit. Allowed claim: local per-part `passthrough`/`regenerate`/`add`/`drop`/`error` evidence. Forbidden: chart byte-passthrough, Excel recalculation equivalence, SLSA, in-toto, signed provenance, third-party attestation, or full streaming parity.
```
