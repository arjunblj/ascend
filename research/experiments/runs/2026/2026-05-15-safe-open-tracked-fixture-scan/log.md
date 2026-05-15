# Safe Open Tracked Fixture Scan

## Question

Should the safe-open fixture replacement scan use only tracked public fixtures when deciding whether generated signed and unknown-part packages can be replaced?

## Hypothesis

Yes. The release blocker is about publishable proof evidence, so ignored local fixture folders should not affect the answer. A tracked-corpus scan should still find no signed or unknown-part replacements, but with cleaner provenance.

## External sources checked

- Microsoft Protected View documentation: https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653
- Microsoft OPC overview: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
- OOXML digital signature origin part: https://ooxml.info/docs/15/15.2/15.2.7/
- OOXML unknown relationships: https://ooxml.info/docs/9/9.1/9.1.7/

## Why this matters to Ascend

Safe unknown workbook opening is the top release claim. Its fixture blocker should be based on release-eligible evidence, not whatever ignored external or stress fixtures happen to exist locally.

## Probe/implementation

Updated `fixtures/benchmarks/safe-open-fixture-scan.ts` to use `git ls-files` by default, falling back to a filesystem walk only when git metadata is unavailable. The fallback skips ignored `external` and `stress` fixture folders. The scan result now records `corpus` and `skippedDirectories`.

## Results

`bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json` now reports:

| Metric | Result |
| --- | ---: |
| Corpus | tracked git fixtures |
| Scanned XLSX/XLSM fixtures | 223 |
| Rejected fixtures | 1 |
| Signed/unknown matches | 0 |

The replacement status remains `no-public-binary-replacement-found`. This strengthens the blocker by proving it over tracked fixtures only; it does not close the blocker.

Validation:

```bash
bun test fixtures/benchmarks/safe-open-fixture-scan.test.ts
bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json
```

## Confidence

High for the tracked-corpus result. Medium for broader public availability: the scan does not prove that no license-clear signed or unknown-part public workbook exists elsewhere.

## Fold-in decision

Promote to product/release proof packaging. Keep `public-edge-fixtures` missing because signed and unknown-part replacements are still absent from the tracked corpus.

## Next question

Should product accept disclosed generated signed/unknown structural packages for guarded safe-open proof, or should a future owner define an external fixture acquisition policy?
