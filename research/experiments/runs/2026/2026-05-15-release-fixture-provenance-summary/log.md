# Release Fixture Provenance Summary

## Question

Can the release proof index distinguish public binary fixture evidence from generated edge-package evidence without checking in large workbook artifacts?

## Hypothesis

Yes. The proof harnesses already know enough about each case to expose a compact provenance summary. The release index should publish counts and case names for public fixtures, generated workbooks, generated edge packages, and malformed inputs.

## External sources checked

- [SLSA provenance](https://slsa.dev/spec/v0.1/provenance) frames provenance as identifying produced artifacts and the input materials used to produce them.
- [GitHub artifact attestation provenance docs](https://docs.github.com/actions/security-for-github-actions/using-artifact-attestations/using-artifact-attestations-to-establish-provenance-for-builds) require subject digests when establishing build provenance.
- [GitHub CLI attestation verification](https://cli.github.com/manual/gh_attestation_verify) describes an attestation as a claim about a subject artifact. Ascend's release index remains below that bar, so source labels must be explicit.

## Why this matters to Ascend

The top release proof blockers say some evidence is generated rather than public binary fixture evidence. That was previously only prose. A machine-readable fixture provenance summary makes the blocker auditable and prevents stronger release copy from hiding synthetic edge cases.

## Probe/implementation

- Added `sourceKind` to package-action proof cases and results:
  - `public-fixture`
  - `generated-workbook`
  - `generated-edge-package`
- Added `fixtureProvenance` to each `release-proof-index` artifact:
  - public fixture case count and names,
  - generated workbook case count,
  - generated edge package case count,
  - malformed case count,
  - generated case names,
  - boundary text.
- Rendered fixture provenance in the Markdown release proof index table.

## Results

Validation passed:

```bash
bun test fixtures/benchmarks/release-proof-index.test.ts fixtures/benchmarks/package-action-proof.test.ts fixtures/benchmarks/safe-open-proof.test.ts
bunx biome check fixtures/benchmarks/release-proof-index.ts fixtures/benchmarks/release-proof-index.test.ts fixtures/benchmarks/package-action-proof.ts fixtures/benchmarks/package-action-proof.test.ts
bunx tsc --build
```

Current no-timings release index now reports:

- `safe-open-proof`: 6 public fixture cases, 0 generated workbook cases, 2 generated edge-package cases, 1 malformed case.
- `package-action-proof`: 2 public fixture cases, 2 generated workbook cases, 4 generated edge-package cases, 0 malformed cases.

## Confidence

High for release-index evidence hygiene. Medium for external publication readiness because generated edge packages still need either explicit product acceptance or replacement by public binary fixtures.

## Fold-in decision

Promote to release/product loop as proof-index metadata. Do not add large fixture binaries and do not claim signed provenance.

Allowed wording: "release proof artifacts now disclose public versus generated fixture provenance by case count and name."

## Next question

Should generated edge-package fixtures be serialized as small deterministic byte digests in the proof output, or is case-name/source-kind disclosure enough until product chooses a publication policy?
