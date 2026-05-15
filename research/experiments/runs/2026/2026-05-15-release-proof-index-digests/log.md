# Release Proof Index Digests

## Question

Can Ascend generate a single release proof index that references safe-open and package-action harness outputs by digest, without embedding bulky artifacts or implying signed provenance?

## Hypothesis

Yes. A benchmark-only proof index can run the existing harnesses, hash their JSON artifacts, record stable shape digests, and render a small Markdown/JSON index that is explicitly not SLSA, in-toto, signed provenance, or tamper-evident storage.

## External sources checked

- SLSA provenance specification: https://slsa.dev/spec/v0.2/provenance
- in-toto attestation framework repository: https://github.com/in-toto/attestation
- GitHub artifact attestations describe subjects bound to digests in in-toto format: https://github.com/actions/attest
- Open Packaging Conventions digital signatures confirm signed package contents have not changed since signing: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/digital-signatures

## Why this matters to Ascend

The release claim board says safe-open and package-action proof harnesses should remain sibling release evidence artifacts, referenced by digest later rather than embedded into every workbook-level release proof bundle. A digest index is the smallest concrete artifact that proves this packaging strategy without adding SDK, CLI, API, or MCP surface area.

## Probe/implementation

- Inspected existing proof harnesses:
  - `fixtures/benchmarks/safe-open-proof.ts`
  - `fixtures/benchmarks/package-action-proof.ts`
- Added `fixtures/benchmarks/release-proof-index.ts`.
  - Runs the two proof harnesses.
  - Emits a compact Markdown or JSON evidence index.
  - Records each artifact command, SHA-256 of the JSON artifact, stable shape SHA-256 with run noise stripped, JSON bytes, Markdown bytes, compact summary, and boundary.
  - Sets `signed: false` and `attestation: false`.
- Added `fixtures/benchmarks/release-proof-index.test.ts`.
  - Asserts both top-claim artifacts are indexed.
  - Asserts SHA-256 formatting.
  - Asserts stable shape digests are deterministic in no-timings mode.
  - Asserts non-attestation wording appears in the Markdown report.

## Results

Local proof command:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings
```

Current no-timings index:

- `safe-open-proof`: 9 cases, 8 ok, 1 rejected, 4 review-before-hydration cases, malformed rejected.
- `package-action-proof`: 8 cases, all five action kinds covered, source graph evidence everywhere.
- The report records artifact digests and byte sizes but does not embed the full proof artifacts.
- The report explicitly states it is not signed provenance, SLSA, in-toto attestation, or tamper-evident storage.

Validation passed:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts
bun run fixtures/benchmarks/release-proof-index.ts --no-timings
```

## Confidence

High that a digest-only index is the right next proof artifact boundary. Medium that the exact schema should be productized later; this should stay a benchmark/proof harness until artifact storage, privacy, and verification semantics are designed.

## Fold-in decision

Fold into research/benchmark proof tooling only. Do not promote to SDK, CLI, API, or MCP. Promote to topic synthesis as evidence that top-claim proof artifacts can be packaged by digest without fake attestation claims.

## Next question

Can the token-bounded agent view claim be tied to a similar digest-index proof artifact, or should it remain below the top two until product defines recovery flows for omitted evidence?
