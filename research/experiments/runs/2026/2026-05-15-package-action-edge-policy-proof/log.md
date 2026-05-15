# Package Action Edge Policy Proof

## Question

Can the package-action `edge-fixture-policy` blocker be resolved by accepting disclosed generated structural edge packages, while still forbidding provenance, trust, and full streaming-parity wording?

## Hypothesis

Yes for guarded structural package-action wording. The compact report already discloses source kinds and keeps headline wording blocked. Product can accept generated edge packages as local structural proof only if release and correctness boundaries remain explicit.

## External sources checked

- Open Packaging Conventions define packages as parts and relationships, which is the structural layer under this proof: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- SLSA provenance describes signed build provenance expectations that the local proof explicitly does not satisfy: https://slsa.dev/spec/v1.0-rc1/provenance
- SheetJS write options document writer/preservation scope for competitor contrast: https://docs.sheetjs.com/docs/api/write-options/
- openpyxl documents preservation boundaries for unsupported workbook objects: https://openpyxl.readthedocs.io/en/stable/tutorial.html

## Why this matters to Ascend

Auditable package-part mutation is the second top product-shaped handoff. The proof is already strong structurally, but the remaining blocker is claim wording: generated edge packages are useful if disclosed, dangerous if they imply real-world provenance or semantic support.

## Probe/implementation

Ran the compact package-action proof:

```bash
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
```

No production code changed.

## Results

- `headlineClaimAllowed=false`
- `releaseGate=blocked-by-publication-policy`
- Source case counts:
  - public fixture: 2
  - generated workbook: 2
  - generated edge package: 4
- Combined commit action counts:
  - `passthrough=27`
  - `regenerate=38`
  - `add=3`
  - `drop=3`
  - `error=1`
- Coverage:
  - 8 cases
  - expected actions everywhere
  - source graph everywhere
  - package journal issues everywhere
  - post-write audit failure only for `unknown-part-error`
  - one representative streaming proof case
- Remaining `readyWhen` blockers: `edge-fixture-policy`, `provenance-boundary`, `unsupported-feature-boundary`, and `streaming-matrix-boundary`.

## Confidence

High for structural package-action coverage. Medium for release wording until product, release, correctness, and performance owners explicitly accept the boundaries.

## Fold-in decision

Promote to topic synthesis. Do not add mutation surfaces. Recommended owner decision: accept generated edge packages only as disclosed local structural proof, while keeping headline claims blocked until provenance, unsupported-feature, and streaming wording are approved.

## Next question

Can release approve a compact-report publication policy that stores no workbook bytes and treats compact report commands as reproducibility pointers rather than published proof artifacts?
