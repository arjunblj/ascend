# Release Proof Index Refresh

## Question

Can the release proof index still point to the refreshed safe-open and package-action proof artifacts by digest without embedding bulky JSON or implying tamper-evident attestation?

## Hypothesis

Yes. The existing benchmark-only release proof index should remain the right boundary: it references the two top-claim artifacts by digest and stable shape digest, while explicitly setting `signed: false` and `attestation: false`.

## External sources checked

- SLSA provenance specification: https://slsa.dev/spec/latest/
- in-toto Attestation Framework: https://github.com/in-toto/attestation
- GitHub artifact attestations: https://github.com/actions/attest
- GitHub REST artifact attestation verification note: https://docs.github.com/en/rest/users/attestations

## Why this matters to Ascend

The top two release claims are now proof-backed locally. A digest index can tie those reports together for release evidence without turning the proof harness into a production surface or claiming a security property Ascend does not provide.

## Probe/implementation

- Inspected `fixtures/benchmarks/release-proof-index.ts` and `fixtures/benchmarks/release-proof-index.test.ts`.
- Reran:
  - `bun run fixtures/benchmarks/release-proof-index.ts --no-timings`
  - `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json`
  - `bun test fixtures/benchmarks/release-proof-index.test.ts`
- Did not change production code or product surfaces.

## Results

The index still contains only the top two artifacts:

| Artifact | Claim | JSON bytes | Markdown bytes | Stable shape digest | Summary |
| --- | --- | ---: | ---: | --- | --- |
| safe-open-proof | safe unknown workbook opening | 3474 | 1946 | `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178` | cases=9, ok=8, rejected=1, reviewBeforeHydration=4, malformedRejected=true |
| package-action-proof | auditable package-part mutation | 11244 | 2687 | `b9758496346c97920c80ba08b6632315708a6d6cc770927695337e729554dbb0` | cases=8, passthrough=27, regenerate=38, add=3, drop=3, error=1, allActionsCovered=true, sourceGraphEverywhere=true |

The Markdown index says:

- `Signed: false`
- `Attestation: false`
- not signed provenance;
- not SLSA;
- not in-toto attestation;
- not tamper-evident storage.

## Confidence

High that the digest index is the correct proof-packaging boundary today. Medium that this schema should stay unchanged if release artifacts become persisted, because persistent storage would need privacy, retention, and verification semantics.

## Fold-in decision

Promote to topic synthesis/proof packaging only. Keep the index in benchmark proof tooling. Do not add SDK, CLI, API, or MCP surfaces for this index yet.

## Next question

Should token-bounded agent view remain below the top two claims until product defines omitted-evidence recovery flows, or is its current proof strong enough for a product example without release-index promotion?
