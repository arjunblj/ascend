# Top Proof Owner Prompts

## Question

Can the top-two release proof artifacts be summarized as next-loop owner prompts with no additional implementation in this block?

## Hypothesis

Yes. The current release proof index already exposes reproducible artifact commands, stable shape digests, publication blockers, and non-attestation boundaries. The useful fold-in is to make the owner prompts concrete enough that future loops publish or reject the claims instead of adding more surfaces.

## External sources checked

- Microsoft Protected View: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Open Packaging Conventions overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- in-toto Attestation Framework: https://github.com/in-toto/attestation
- SLSA attestation model: https://slsa.dev/spec/v1.1/attestation-model

## Why this matters to Ascend

The ranked portfolio says research should not keep adding narrow surfaces after the claim board identifies the top two release claims. Owner prompts are a control surface for the next loops: they name the evidence, blockers, boundaries, and exit conditions that make the claims publishable or keep them out of release copy.

## Probe/implementation

Ran the release proof index:

```bash
bun run fixtures/benchmarks/release-proof-index.ts --no-timings --json
```

Current artifact facts:

- `safe-open-proof` command: `bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json`
- `safe-open-proof` stable shape digest: `6aa54a651309b3c45ce7ce93ff7034e7b31e47c7cbc458c58ee6a6f23e0c6178`
- `package-action-proof` command: `bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json`
- `package-action-proof` stable shape digest: `b9758496346c97920c80ba08b6632315708a6d6cc770927695337e729554dbb0`
- Both artifacts remain `needs-release-packaging`.

Folded those details into `research/experiments/syntheses/2026-05-release-claim-board.md` without touching SDK, CLI, API, MCP, or writer behavior.

## Results

The release claim board now gives each top owner:

- exact reproduction command;
- stable digest expectation;
- required proof fields;
- current publication blockers;
- explicit "do not claim" boundaries;
- blocking exit criteria.

Safe-open remains blocked from stronger headline copy if code-generated signed/unknown package evidence is hidden or if timing language reads as a threshold.

Package-action remains blocked from stronger headline copy if synthetic edge packages are hidden, if chart XML is called byte-passthrough, or if local proof is framed as SLSA/in-toto/signed provenance.

## Confidence

High for the handoff shape. It is derived from the current release proof index and primary external references. Medium for publication readiness because the blockers are intentionally unresolved.

## Fold-in decision

Promote to topic synthesis only. Do not add production surfaces. Hand off exactly two owner loops: safe unknown workbook opening and auditable package-part mutation.

## Next question

Can the next research cycle prove whether any public binary fixtures can replace the current code-generated signed and unknown package cases, or should those blockers be accepted as disclosed synthetic proof?
