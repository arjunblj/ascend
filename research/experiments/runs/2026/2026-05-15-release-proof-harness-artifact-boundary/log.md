# Release Proof Harness Artifact Boundary

## Question

Should the release proof bundle consume the safe-open and package-action proof harness outputs as embedded local evidence artifacts now?

## Hypothesis

Not yet. The current release proof bundle is per-workbook mutation evidence: inspect, plan, commit, reopen, diff, audit, package actions, and consistency checks for one subject workbook/output pair. The safe-open and package-action harnesses are suite-level claim evidence across multiple fixtures. They should be sibling release evidence artifacts with digests, not embedded into every workbook-level release proof bundle until product defines artifact storage and verification semantics.

## External sources checked

- GitHub artifact attestations establish build provenance for specific artifact subjects and require a real attestation workflow: https://docs.github.com/actions/how-tos/secure-your-work/use-artifact-attestations/use-artifact-attestations
- SLSA provenance defines provenance as a claim that a builder produced subject artifacts from a recipe and materials: https://slsa.dev/spec/v0.1/provenance
- in-toto attestation framework defines verifiable claims about how artifacts are produced and consumed by policy engines: https://github.com/in-toto/attestation
- GitHub CLI attestation verification treats attestations as signed claims about artifact subjects: https://cli.github.com/manual/gh_attestation_verify

## Why this matters to Ascend

Ascend's release-claim board is intentionally conservative. Adding harness outputs into `createReleaseProofBundle()` would make the release proof look broader, but it would mix two different evidence levels:

- workbook-level evidence for one mutation subject;
- suite-level evidence that a release claim holds across representative fixtures.

That mix risks implying signed provenance, full-release certification, or claim coverage that the bundle does not verify for the subject workbook.

## Probe/implementation

- Inspected `ReleaseProofBundle`, `createReleaseProofBundle()`, release proof tests, prior release-proof promotion gate, safe-open proof harness, and package-action proof harness.
- Ran a local probe that created a simple workbook, planned and committed a `setCells` edit, built a release proof bundle with source/output bytes and diff evidence, and compared its JSON size and subject boundary against the package-action harness output.
- Did not edit production code. The worktree currently has unrelated dirty SDK files, and the evidence supports a hold decision rather than another surface.

Probe output:

```json
{
  "releaseProof": {
    "kind": "ascend-release-proof-bundle",
    "proofKind": "local-evidence",
    "jsonBytes": 23733,
    "consistency": true,
    "packageActionIssueCount": 0
  },
  "packageHarness": {
    "caseCount": 8,
    "jsonBytes": 6988,
    "combinedCommitActionCounts": {
      "passthrough": 27,
      "regenerate": 38,
      "add": 3,
      "drop": 3,
      "error": 1
    },
    "hasUnknownReviewCase": true
  }
}
```

The release proof already carries explicit boundaries:

- local workbook evidence, not signed SLSA/GitHub attestation;
- no Excel recalculation equivalence claim;
- no private workbook bytes by default;
- signed provenance requires an external attestation envelope and verifier roots.

## Results

- Current `createReleaseProofBundle()` is correctly scoped for one workbook mutation and already carries package-action digest evidence when source/output bytes are supplied.
- Safe-open and package-action proof harnesses should be generated as separate release evidence artifacts, then referenced by digest from release notes or a future release evidence manifest.
- Embedding suite-level harness output into every workbook mutation proof would blur subject semantics and increase proof size without improving per-workbook correctness.
- A future product loop can add a release evidence manifest that lists artifact names, digests, commands, fixture scope, and claim boundaries.

Validation:

- Local probe executed successfully.
- Markdown-only cycle; no heavy suite needed.

## Confidence

High for the hold decision. It matches SLSA/in-toto/GitHub attestation subject boundaries and Ascend's existing non-attestation wording. Medium for the exact future manifest shape; product should decide whether the manifest is a CLI-generated JSON file, release-note table, or CI artifact.

## Fold-in decision

Promote to topic synthesis / product handoff, not production. Do not embed harness outputs in `ReleaseProofBundle` yet. The next fold-in should be a release evidence manifest or CLI artifact only after product defines storage, privacy, and verification semantics.

## Next question

Can a token-bounded agent-view proof harness show deterministic budget adherence and omitted-evidence recovery across mixed workbook shapes without changing agent-view surfaces?
