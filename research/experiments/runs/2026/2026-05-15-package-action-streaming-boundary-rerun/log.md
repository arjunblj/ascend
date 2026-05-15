# Package Action Streaming Boundary Rerun

## Question

Can the package-action `streaming-matrix-boundary` blocker be closed by current evidence, or does it still require performance-owner approval or a broader streaming matrix?

## Hypothesis

It cannot be closed by current evidence. The package-action proof still has one representative streaming dirty-sheet proof, while add/drop/error and public macro/chart cases remain standard writer proof only. That is useful, but it is not full streaming parity.

## External sources checked

- Open Packaging Conventions describe the package part and relationship layer whose preservation is being measured: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- SheetJS CE documents format read/write scope and reinforces that spreadsheet writer support varies by feature: https://docs.sheetjs.com/docs/api/write-options/
- Excelize documents a dedicated stream writer, useful competitor context for separating streaming write claims from general XLSX write claims: https://xuri.me/excelize/en/stream.html

## Why this matters to Ascend

The auditable package-part mutation claim is credible today for the standard writer proof. If release wording drifts into streaming parity, it would overclaim: streaming proof currently covers a representative dirty-sheet passthrough/regenerate case, not the full action matrix.

## Probe/implementation

Ran:

```bash
bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json
rg -n "writeXlsxStreaming|streamingProbe|streaming-matrix" packages fixtures research -g '*.ts' -g '*.md'
```

No production code changed. Existing harness tests already assert `streamingProofCases: 1`, `streamingRegenerateParts: 1`, and compact boundary text saying this is not full streaming parity.

## Results

Current compact proof:

- 8 package-action cases;
- all five action classes covered in standard commit proof: `passthrough=27`, `regenerate=38`, `add=3`, `drop=3`, `error=1`;
- source graph and journal package issues are present in every case;
- `unknown-part-error` is the only post-write audit failure;
- streaming proof cases: 1;
- streaming regenerated parts: 1;
- the only streaming case is `docprops-passthrough`, with one worksheet regeneration and three byte-equal passthrough parts.

The existing source scan confirms `streamingProbe: true` is assigned only to `docprops-passthrough`.

## Confidence

High that current evidence supports only narrow streaming wording. Medium that expanding to every package-action scenario is worth doing; add/drop/error and macro/chart cases may require deliberate streaming writer design work rather than a benchmark-only toggle.

## Fold-in decision

Archive as proof-boundary rerun. Keep `streaming-matrix-boundary` missing. The allowed wording remains: "one representative streaming dirty-sheet package-action proof exists." Do not say streaming parity covers add/drop/error, macro/chart public fixtures, or every package-action scenario.

## Next question

Should performance owners accept the narrow representative streaming wording for release, or explicitly fund a broader streaming writer matrix that covers add/drop/error and public macro/chart fixtures?
