# Package Action Proof

## Question

Can package writes be fully explained as `passthrough`, `regenerate`, `add`, `drop`, or `error` with proof artifacts?

## Hypothesis

Mostly yes. Ascend already has the right raw evidence: write-plan origins, package graph parts, preservation policies, write-policy diagnostics, package graph audits, trace digests, and post-write verification. The missing layer is a per-package-part action taxonomy that translates origin and policy into a proof shape an agent can audit without knowing writer internals.

## External sources checked

- [Microsoft Open Packaging Conventions fundamentals](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview): packages are directed graphs of URI-addressable parts and relationships, and digital signatures can reference package components.
- [Microsoft Open XML SDK packages and general docs](https://learn.microsoft.com/en-us/office/open-xml/general/overview): package work is organized around document parts and relationships.
- [Microsoft `PackageDigitalSignatureManager`](https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging.packagedigitalsignaturemanager?view=windowsdesktop-10.0): signatures are first-class package metadata, supporting Ascend's current invalidate-on-edit treatment.
- [SheetJS parse options](https://docs.sheetjs.com/docs/api/parse-options): `bookVBA` exposes raw VBA blobs and metadata-only parse modes, but does not present a write proof taxonomy.
- [openpyxl tutorial](https://openpyxl.readthedocs.io/en/stable/tutorial.html): `keep_vba` is explicit and the docs warn that unsupported shapes are lost, reinforcing that preservation claims need auditable detail.

## Why this matters to Ascend

This is a trust differentiator for preservation-first XLSX. Existing libraries often expose capability caveats as options or warnings. Ascend can make a stronger claim: for each package part, say whether it was passed through, regenerated, added, intentionally dropped, or blocked, with the evidence that justified the action.

## Probe/implementation

Inspected local implementation:

- `packages/io-xlsx/src/writer/plan.ts` defines `WritePartOrigin = generated | preserved-inline | preserved-source | capsule` and emits part summaries.
- `packages/io-xlsx/src/package-graph.ts` classifies package parts, feature families, owner scopes, preservation policies, and byte-preservation expectations.
- `packages/sdk/src/agent-workflow.ts` combines preservation, package graph audit, write-policy diagnostics, approvals, and trace artifacts.
- `packages/sdk/src/agent-workflow.test.ts` already exercises inspect-only, review-required, signature invalidation, calc-chain, visual, table, comment, and external-link policy paths.

Added an ignored local probe at `research/experiments/runs/2026/2026-05-14-package-action-proof/probes/action-proof-candidate.ts`. It opens three local workbooks, plans one `setCells` edit, compares planned write parts against `inspectXlsxPackageGraph(sourceBytes)`, and derives candidate actions:

- `generated` + source part present -> `regenerate`
- `generated` + source part absent -> `add`
- `preserved-source`, `preserved-inline`, or `capsule` -> `passthrough`
- diagnostics with `discarded-for-recalc` or `invalidated-on-edit` -> candidate `drop`
- blocker diagnostics -> candidate `error`

Validation command:

```bash
bun run research/experiments/runs/2026/2026-05-14-package-action-proof/probes/action-proof-candidate.ts
```

## Results

The probe successfully derived action summaries without production changes:

| Workbook | Source parts | Planned parts | Passthrough | Regenerate | Add | Drop | Error | Notes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| `research/real-workbooks/Book1.xlsx` | 9 | 8 | 2 | 6 | 0 | 0 | 0 | clean edit, package graph audit passed |
| `research/excel-corpus/conditional-formatting.xlsx` | 17 | 16 | 10 | 6 | 0 | 0 | 0 | approval required for preservation policy, but actions still inferable |
| `research/excel-corpus/bevreport-demo.xlsm` | 318 | 317 | 302 | 15 | 0 | 0 | 0 | macro/visual/table warnings surfaced separately from action counts |

The taxonomy works as a candidate overlay, but two gaps remain:

- `passthrough` lacks byte-hash proof in the compact plan. The writer knows bytes are from source/capsule, but the proof bundle should include source and output hashes per passed-through part.
- `drop` and `error` should not be inferred only from diagnostics. They need explicit action records with paths, reasons, triggering operation, and expected post-write audit status.

## Confidence

High that `passthrough/regenerate/add` can be layered over current write-plan and package graph data. Medium for `drop/error`, because they need a production-facing proof model rather than inference from warning text.

## Fold-in decision

Promote to the correctness loop. Define a small `PackageDeltaPlan` or `PackageActionProof` surface that wraps existing plan data without changing writer behavior first.

Promote to the product/DX loop after correctness owns the shape. CLI/API/MCP should expose compact counts plus sampled action records so agents can decide whether a write is trustworthy.

Promote to the performance loop later. The same action proof can guide mixed passthrough/regenerate writes and prevent accidental full rewrites.

## Next question

Can formula AST spans provide enough structure for hover, diagnostics, rename, and safe code actions without changing formula evaluation semantics?
