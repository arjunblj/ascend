# Package Proof Compact JSON Gate

## Question

Should Ascend track a stable compact JSON fixture for the package-action proof report now?

## Hypothesis

No. The compact JSON is useful as generated evidence, but the current proof is already covered by a tracked harness and Markdown report. Tracking generated JSON would add another artifact format before product has defined publication, privacy, and retention rules.

## External sources checked

- in-toto statements bind attestations to subjects and predicate types; Ascend's package proof is not yet an attestation: https://github.com/in-toto/attestation/blob/main/spec/README.md
- Docker's SLSA provenance docs distinguish minimal versus maximal provenance and call out leakage risk from richer build metadata: https://docs.docker.com/build/metadata/attestations/slsa-provenance/
- Docker's SLSA definitions note that local untracked source has incomplete materials compared with tracked remote source: https://docs.docker.com/build/metadata/attestations/slsa-definitions/
- OpenSSF Scorecard is a repository-level signal, not a substitute for product-specific proof artifacts: https://openssf.org/scorecard/

## Why this matters to Ascend

The auditable package-part mutation claim needs durable proof, but premature artifacts can create false authority. A JSON fixture should be tracked only when it is stable, intentionally scoped, and aligned with a product proof bundle format.

## Probe/implementation

Ran a local generated-artifact size probe against the current package-action proof harness:

```bash
bun --eval "import { runPackageActionProof } from './fixtures/benchmarks/package-action-proof.ts'; const result = await runPackageActionProof({ includeTimings: false }); ..."
```

Probe result:

- cases: `8`
- full generated JSON bytes: `7854`
- compact generated JSON bytes: `1896`
- every case had `commitJournalPackageIssueCount = 1`

No production code was changed. No JSON artifact was added.

## Results

The compact JSON shape is small enough to publish later, but not worth tracking yet:

- It would duplicate the tracked harness assertions.
- It would need generated timestamps removed before hash stability.
- `research/**/*.json` is intentionally ignored to keep generated artifacts out of git.
- The release proof board still says publication packaging and privacy semantics are unresolved.

## Confidence

High for not promoting a tracked JSON fixture in this block. Medium for future JSON publication after product defines a stable release-proof artifact schema.

## Fold-in decision

Archive as "do not promote yet." Keep the tracked TypeScript harness as the validation gate, and keep Markdown reports as the claim-steward artifact.

## Next question

Can retained viewport patch history be rerun as a claim-steward proof report without adding CLI/API/MCP surfaces?
