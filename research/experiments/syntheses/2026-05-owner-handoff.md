# 2026-05 Owner Handoff

Date: 2026-05-15

## Decision

Hand off only two implementation loops:

1. **Safe unknown workbook opening** to product/performance.
2. **Auditable package-part mutation** to correctness/product, with release and performance boundary approvals.

Everything else remains do-not-promote for this block. Formula intelligence stays rejection-first; do not implement rename.

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
| Public fixture cases | 2 |
| Generated workbook cases | 2 |
| Generated edge-package cases | 4 |
| Generated/disclosed cases | `docprops-passthrough`, `regenerate-existing-sheet`, `add-sheet-part`, `calc-chain-drop`, `signature-invalidation-drop`, `unknown-part-error` |
| Action classes | `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1` |
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

## Do Not Promote

| Direction | Freeze reason |
| --- | --- |
| Formula language-service primitives | Rejection-first proof exists, but rename remains unsafe without workbook-context symbol ownership and operation-owned edits. |
| Retained viewport patch history | Product-proof backed, but not a top release handoff and not collaboration/CRDT evidence. |
| Token-bounded agent view | Product-proof backed, but token counts are approximate and omitted evidence is by design. |
| Release proof bundle | Ingredients exist, but artifact storage/privacy/canonicalization and non-attestation policy are unresolved. |
| Formula conformance/oracle routing | Still a correctness research program; no broad Excel-compatible formula claim. |
| Property-style journal laws | Useful correctness evidence, but not a release claim until coverage and shrinkability are owner-approved. |
| Columnar scan sidecars | Performance research only; not a product surface. |
| Agent workflow observability | Supporting DX evidence, not a top release claim. |

## Next-Loop Prompts

### Product/Performance

```text
/goal Become Ascend's safe unknown workbook opening proof owner. Do not add new open surfaces. Start from `fixtures/benchmarks/release-proof-index.ts` and `safe-open-proof`. Resolve or explicitly accept generated signed/unknown package fixtures, run release-environment open-plan latency on standardized public inputs, and approve wording that says pre-hydration package-feature routing and review recommendation only. Do not claim malware scanning, sandboxing, trust, active-content safety, signed provenance, or release latency thresholds. Commit only proof/report/index changes or narrow harness fixes.
```

### Correctness/Product

```text
/goal Become Ascend's auditable package-part mutation proof owner. Do not add mutation surfaces. Start from `fixtures/benchmarks/release-proof-index.ts` and `package-action-proof`. Resolve generated edge-package policy, approve the unsupported-feature matrix, preserve journal/package issue compatibility, and keep streaming/provenance boundaries explicit. Allowed claim: local per-part `passthrough`/`regenerate`/`add`/`drop`/`error` evidence. Forbidden: chart byte-passthrough, Excel recalculation equivalence, SLSA, in-toto, signed provenance, third-party attestation, or full streaming parity.
```
