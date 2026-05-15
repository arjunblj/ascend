# Release Proof Index Current Refresh

Date: 2026-05-15

## Question

After refreshing the safe-open and package-action proof packages, can the release proof index still point only to those top two artifacts by digest without implying signed provenance, SLSA, in-toto, or tamper-evident storage?

## Hypothesis

Yes. The release index should stay a benchmark/proof artifact: it references local evidence by digest and stable shape digest, sets `signed: false` and `attestation: false`, and does not promote formula, agent-view, or viewport artifacts into release claims.

## External Sources Checked

- in-toto describes attestations as verifiable claims about how software is produced: https://github.com/in-toto/attestation
- SLSA provenance describes verifiable artifact production metadata and stricter integrity requirements: https://slsa.dev/spec/v1.0/provenance
- Open Packaging Conventions keep the XLSX proof boundary at package parts/relationships, not release attestation: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview

## Why This Matters To Ascend

The claim stewarding block should hand off the top two product claims with proof, not keep expanding release scope. A small digest index is enough to bind the current safe-open and package-action proof artifacts together while preserving the honest boundary that this is local evidence only.

## Probe/Implementation

No production code changed. Reran the existing proof index:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings
bun test fixtures/benchmarks/release-proof-index.test.ts
```

## Results

Latest proof index run: 2026-05-15T03:58:23.849Z.

| Artifact | Claim | JSON bytes | Markdown bytes | SHA-256 | Stable shape SHA-256 | Summary |
| --- | --- | ---: | ---: | --- | --- | --- |
| safe-open-proof | safe unknown workbook opening | 3474 | 1946 | `a5d3d59e8bc0b510b45ccbd4ff1ac102f0cc97551f419c816041cbd0fa6287ab` | `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178` | cases=9, ok=8, rejected=1, reviewBeforeHydration=4, malformedRejected=true |
| package-action-proof | auditable package-part mutation | 11244 | 2687 | `7b75985b96c5ed9060566785188f2a91859785f5567e8897a7eb3eff25252b28` | `b9758496346c97920c80ba08b6632315708a6d6cc770927695337e729554dbb0` | cases=8, passthrough=27, regenerate=38, add=3, drop=3, error=1, allActionsCovered=true, sourceGraphEverywhere=true |

The rendered index states:

- `Signed: false`
- `Attestation: false`
- not signed provenance
- not SLSA
- not in-toto attestation
- not tamper-evident storage

## Confidence

High that the release proof index is correctly scoped today. Medium that the raw run SHA should be used for anything durable because it includes generated timestamps; the stable shape digest is the better comparison key for current no-timings artifact shape.

## Fold-In Decision

Promote to product/release handoff only. Do not add SDK, CLI, API, or MCP surfaces for the release proof index. Keep release membership limited to the safe-open and package-action artifacts.

## Next Question

Hand off the top two implementation-loop prompts now: safe-open publication packaging and auditable package-action publication packaging. Do not promote formula rename, agent-view, viewport, or sidecar work into release proof scope yet.
