# Package Action Release Proof Dirty Writer Blocker

## Question

Can auditable package-part mutation receive the same current release-proof rerun as safe-open without changing writer behavior?

## Hypothesis

Not safely in this worktree right now. The tracked proof harness depends on the XLSX writer, and `packages/io-xlsx/src/writer/index.ts` is already dirty with unrelated production changes. A rerun would measure and validate uncommitted writer behavior, which is too weak for a release-claim report over `main`.

## External sources checked

- Microsoft Open Packaging Conventions fundamentals: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- Microsoft `System.IO.Packaging` namespace, including package parts, relationships, and signatures: https://learn.microsoft.com/en-us/dotnet/api/system.io.packaging
- SheetJS write options and unsupported-feature boundary: https://docs.sheetjs.com/docs/api/write-options/
- openpyxl preservation warning: https://openpyxl.readthedocs.io/en/stable/tutorial.html
- in-toto attestation framework boundary for real signed provenance: https://github.com/in-toto/attestation

## Why this matters to Ascend

Auditable package-part mutation is the second-ranked release claim. It must be stricter than a normal research note because the proof report says what happened to package parts during writes. If the writer is dirty from another loop, the current harness can still run, but the evidence would not cleanly describe committed `main` behavior.

## Probe/implementation

- Inspected current worktree status.
- Inspected `fixtures/benchmarks/package-action-proof.ts` and `fixtures/benchmarks/package-action-proof.test.ts`.
- Inspected the unrelated dirty writer diff enough to confirm it changes write planning logic around dirty sheet names and preserved sheet XML patching.
- Did not run or update the package-action release report, because the harness would depend on the dirty writer behavior.
- Did not stage, edit, or revert the dirty writer file.

## Results

Current blocker:

```text
M packages/io-xlsx/src/writer/index.ts
```

The proof harness remains the right artifact after the writer worktree is clean:

```bash
bun run fixtures/benchmarks/package-action-proof.ts
bun test fixtures/benchmarks/package-action-proof.test.ts
```

Expected release-proof scope after unblock:

- `passthrough`: docProps and macro sidecars;
- `regenerate`: sheet XML and chart XML where applicable;
- `add`: new sheet parts;
- `drop`: calc-chain and digital-signature invalidation;
- `error`: unknown package parts requiring review;
- boundary: local package evidence only, not signed provenance, SLSA, in-toto, or Excel recalculation equivalence.

## Confidence

High that this is a real proof-quality blocker. Running the harness now could still be useful for the writer loop, but it would not be clean release-claim evidence.

## Fold-in decision

Archive as blocked for this loop until the writer worktree is clean or the writer owner commits their change. Do not promote a package-action release-proof refresh from this dirty state.

## Next question

Can the next non-writer claim, formula language-service rejection evidence, be tightened without implementing rename or depending on XLSX writer behavior?
