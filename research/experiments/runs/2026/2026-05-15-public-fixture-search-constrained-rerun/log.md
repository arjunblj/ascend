# Public Fixture Search Constrained Rerun

## Question

Can public fixture search be constrained to only the missing signed/unknown structural cases, avoiding another broad corpus sweep?

## Hypothesis

Yes. The useful diagnostic is to rerun the checked-in fixture scanner and do a targeted web check for signed/unknown XLSX fixtures. If neither finds a durable public binary, product approval of disclosed generated structural fixtures remains the practical path.

## External sources checked

- Targeted GitHub search for signed XLSX fixtures containing `_xmlsignatures/origin.sigs`
- Targeted GitHub search for `.xlsx` fixtures with `_xmlsignatures`
- Targeted GitHub search for OOXML unknown relationship XLSX fixtures
- Targeted GitHub search for signed XLSX test fixtures

Notable search results were library repositories or issues discussing unknown relationships/macros, not durable public signed/unknown XLSX binaries suitable to vendor.

## Why this matters to Ascend

Safe unknown workbook opening is blocked on public fixture evidence or explicit acceptance of generated structural fixtures. Another broad corpus sweep would waste time unless it is constrained to the missing edge cases.

## Probe/implementation

Ran:

```bash
bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json
```

## Results

| Field | Value |
| --- | --- |
| Fixture root | `fixtures/xlsx` |
| Scanned fixtures | 351 |
| Rejected fixtures | 2 |
| Rejected paths | `fixtures/xlsx/calamine/pass_protected.xlsx`, `fixtures/xlsx/poi/protected_passtika.xlsx` |
| Signature/unknown matches | 0 |
| Replacement status | `no-public-binary-replacement-found` |

Boundary: this scans checked-in public XLSX/XLSM fixtures only. It does not prove that no public signed or unknown-part workbooks exist elsewhere.

## Confidence

High that the current repo has no replacement. Medium that a broader public web search will remain unproductive; targeted search did not find a clean fixture candidate, but the web is not exhaustive.

## Fold-in decision

Promote to topic synthesis only. Keep `safe-open-proof/public-edge-fixtures` missing until product accepts generated structural fixtures or provides public binaries.

## Next question

Can compact report privacy review be reduced to field-level inventory rather than adding canonicalization or storage policy?
