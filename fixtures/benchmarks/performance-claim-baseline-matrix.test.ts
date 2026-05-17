import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'

const REPO_ROOT = new URL('../../', import.meta.url)
const MATRIX_PATH = 'docs/PERFORMANCE_CLAIM_BASELINE_MATRIX.md'

const REQUIRED_WORKLOAD_GAPS = [
	'dense-values',
	'sparse-wide',
	'styles-heavy',
	'formula-heavy',
	'table-heavy',
	'feature-rich',
	'selected-sheet',
	'metadata-only',
	'warm-workflow',
]

const UNAVAILABLE_RUNNER_ROWS = [
	'fastexcel Python',
	'python-calamine',
	'Polars calamine',
	'Polars xlsx2csv',
	'Polars openpyxl',
	'pyopenxlsx',
	'fastxlsx',
]

const RECORDED_WORKLOADS = [
	'dense-values',
	'sparse-wide',
	'styles-heavy',
	'formula-heavy',
	'table-heavy',
	'feature-rich',
	'selected-sheet',
	'metadata-only',
	'warm-workflow',
	'string-heavy',
]

describe('performance claim baseline matrix', () => {
	test('pins scoped read evidence as a defer decision, not a speed claim', () => {
		const markdown = readRepoFile(MATRIX_PATH)

		expect(markdown).toContain('Status: defer.')
		expect(markdown).toContain('This document is the tracked release claim artifact')
		expect(markdown).toContain(
			'No broad XLSX read, XLSX write, SOTA, or QSS-leapfrog speed claim is promotable',
		)
		expect(markdown).toContain(
			'The current full-profile run at `9ddfff91` reports no leader failures',
		)
		expect(markdown).toContain(
			'Current harness evidence now supports same-lane selected-sheet rows for Ascend, SheetJS, OpenPyXL, and python-calamine.',
		)
		expect(markdown).toContain(
			'Current harness evidence now supports same-lane metadata-only rows for Ascend, SheetJS, OpenPyXL, and python-calamine.',
		)
		expect(markdown).toContain('Calamine wins that head-to-head')
		expect(markdown).toContain(
			'Current harness evidence now supports a SheetJS feature-rich rich-metadata row',
		)
		expect(markdown).toContain(
			'Current focused `plain-text` and `string-heavy` write coverage proves\n  ClosedXML and NPOI now run and pass validation',
		)
		expect(markdown).toContain(
			"Current focused TS/JS/Rust `dense-values` write coverage proves Ascend's\n  generated writer is faster by median and p95 than SheetJS, ExcelJS, and\n  rust_xlsxwriter",
		)
		expect(markdown).toContain(
			'Current focused TS/JS/Rust `sparse-wide` write coverage supersedes the older\n  sparse-wide p95 boundary',
		)
		expect(markdown).toContain(
			"Current focused TS/JS/Rust `plain-text` write coverage proves Ascend's\n  generated writer is faster by median and p95 than SheetJS, ExcelJS, and\n  rust_xlsxwriter",
		)
		expect(markdown).toContain(
			'Current formula/calc evidence includes focused HyperFormula indexed\n  `INDEX/MATCH`, indexed dirty-key/dirty-value edits, and prefix-range\n  dirty-head/dirty-tail rows.',
		)
		for (const workload of RECORDED_WORKLOADS) expect(markdown).toContain(`\`${workload}\``)
		expect(markdown).toContain('Humble allowed wording:')
		expect(markdown).toContain('Forbidden wording:')
		expect(markdown).toContain('"Ascend is the fastest XLSX reader."')
		expect(markdown).toContain('"Ascend is SOTA for XLSX read."')
		expect(markdown).toContain('Any wording that treats failed or unavailable runners as wins.')

		expect(markdown).toContain(
			'## Cycle: Formula SOTA Indexed Lookup HyperFormula Row at `cd1c0415`',
		)
		expect(markdown).toContain('Classification: comparable formula-engine evidence plus defer.')
		expect(markdown).toContain('`hf-indexed-index-match`')
		expect(markdown).toContain('commit\n`cd1c0415`')
		expect(markdown).toContain(
			'/private/tmp/ascend-formula-sota-current-cd1c0415-runs/hf-indexed-index-match-repeat30.json',
		)
		expect(markdown).toContain(
			'| Ascend | ran/won | 6.305 ms / 12.225 ms / 0.398 | 29.644 ms / 58.706 ms / 0.379 | 36.690 ms / 67.769 ms / 0.341 |',
		)
		expect(markdown).toContain(
			'| HyperFormula | ran/lost vs Ascend | 1075.377 ms / 1736.006 ms / 0.295 | 437.929 ms / 1031.233 ms / 0.471 | 1539.467 ms / 2709.070 ms / 0.321 |',
		)
		expect(markdown).toContain('`operationSpeedupVsHyperFormula: 14.773x`')
		expect(markdown).toContain(
			'`--profile all --repeat 15 --warmup 3\n--assert-correctness --json` was killed',
		)
		expect(markdown).toContain('"Ascend is SOTA for formula calculation."')
		expect(markdown).toContain(
			'defer production optimization from this winning row. Continue\nformula/calc performance work only with a named HyperFormula workflow',
		)

		expect(markdown).toContain('## Cycle: Formula SOTA Indexed Lookup Dirty Edits at `2700c72a`')
		expect(markdown).toContain('`hf-indexed-index-match-dirty-key`')
		expect(markdown).toContain('`hf-indexed-index-match-dirty-value`')
		expect(markdown).toContain(
			'/private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-key-repeat30.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-formula-indexed-dirty-current-2700c72a-runs/hf-indexed-index-match-dirty-value-repeat30.json',
		)
		expect(markdown).toContain(
			'| Dirty key edit | Ascend | ran/won | 17.453 ms / 23.353 ms / 0.206 | 0.152 ms / 0.552 ms / 1.406 | 17.591 ms / 23.494 ms / 0.212 |',
		)
		expect(markdown).toContain(
			'| Dirty key edit | HyperFormula | ran/lost vs Ascend | 1252.808 ms / 2396.354 ms / 0.401 | 0.659 ms / 2.049 ms / 0.631 | 1254.083 ms / 2397.869 ms / 0.401 |',
		)
		expect(markdown).toContain(
			'| Dirty return-value edit | Ascend | ran/won | 17.903 ms / 27.906 ms / 0.240 | 0.098 ms / 0.221 ms / 1.761 | 18.008 ms / 28.001 ms / 0.237 |',
		)
		expect(markdown).toContain(
			'| Dirty return-value edit | HyperFormula | ran/lost vs Ascend | 1463.408 ms / 3140.009 ms / 0.373 | 347.289 ms / 951.250 ms / 0.448 | 1834.094 ms / 3740.465 ms / 0.354 |',
		)
		expect(markdown).toContain('`operationSpeedupVsHyperFormula: 4.335x`')
		expect(markdown).toContain('`operationSpeedupVsHyperFormula: 3532.498x`')
		expect(markdown).toContain('Ascend operation samples are noisy at sub-millisecond\nscale')

		expect(markdown).toContain(
			'## Cycle: Formula SOTA Prefix Dirty-Tail HyperFormula Row at `c06bba18`',
		)
		expect(markdown).toContain('`hf-prefix-range-dirty-tail`')
		expect(markdown).toContain(
			'/private/tmp/ascend-formula-dirty-current-c06bba18-runs/hf-prefix-range-dirty-tail-repeat30.json',
		)
		expect(markdown).toContain(
			'| Ascend | ran/won | 7.122 ms / 9.880 ms / 0.185 | 0.059 ms / 0.090 ms / 0.208 | 7.184 ms / 9.941 ms / 0.183 |',
		)
		expect(markdown).toContain(
			'| HyperFormula | ran/lost vs Ascend | 53.358 ms / 65.878 ms / 0.171 | 0.091 ms / 0.516 ms / 1.157 | 53.496 ms / 65.946 ms / 0.172 |',
		)
		expect(markdown).toContain('`operationSpeedupVsHyperFormula: 1.531x`')
		expect(markdown).toContain(
			'Changed-cell counts differ\nbecause Ascend reports 1 changed output cell while HyperFormula reports 2',
		)
		expect(markdown).toContain(
			'"Ascend beats HyperFormula on every incremental recalculation workflow."',
		)

		expect(markdown).toContain(
			'## Cycle: Formula SOTA Prefix Dirty-Head HyperFormula Row at `bd91386a`',
		)
		expect(markdown).toContain('`hf-prefix-range-dirty-head`')
		expect(markdown).toContain(
			'/private/tmp/ascend-formula-dirty-head-current-bd91386a-runs/hf-prefix-range-dirty-head-repeat30.json',
		)
		expect(markdown).toContain(
			'| Ascend | ran/won | 6.284 ms / 8.729 ms / 0.372 | 0.810 ms / 0.990 ms / 0.110 | 7.054 ms / 9.612 ms / 0.338 |',
		)
		expect(markdown).toContain(
			'| HyperFormula | ran/lost vs Ascend | 39.465 ms / 46.316 ms / 0.085 | 11.719 ms / 13.636 ms / 0.072 | 51.246 ms / 59.952 ms / 0.073 |',
		)
		expect(markdown).toContain('`operationSpeedupVsHyperFormula: 14.474x`')
		expect(markdown).toContain(
			'Changed-cell counts differ\nbecause Ascend reports 5,000 changed output cells while HyperFormula reports\n5,001',
		)

		expect(markdown).toContain('## Owner-Ready Benchmark Blocker')
		expect(markdown).toContain('Owner: benchmarking/external baselines.')
		expect(markdown).toContain(
			'broad read-speed and QSS-leapfrog performance wording is downgraded',
		)
		expect(markdown).toContain(
			'do not optimize further from the measured winning rows\n`dense-values`, `sparse-wide`, `styles-heavy`, `formula-heavy`, `table-heavy`,\n`selected-sheet`, `metadata-only`, `warm-workflow`, and `string-heavy`',
		)
		expect(markdown).toContain(
			'Failed, missing, or semantically mismatched runners are not counted as wins.',
		)

		expect(markdown).toContain('## Current Full-Profile Downgrade: XLSX Read SOTA')
		expect(markdown).toContain('Classification: claim downgrade.')
		expect(markdown).toContain('Commit: `9ddfff91efc8f0f95edf36f44b78f5313480ad11`')
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-all.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-9ddfff91-runs/xlsx-read-sota-merged-selected-metadata-scoreboard.json',
		)
		expect(markdown).toContain('Current full-profile scoreboard: `leaderFailures: []`')
		expect(markdown).toContain('Merged selected-sheet/metadata-only scoreboard')
		expect(markdown).toContain('`coverageFailures: 10`, `coverageGaps: 8`')
		expect(markdown).toContain(
			'The merged scoreboard removes the selected-sheet and metadata-only\n  `missing-comparable` failures.',
		)
		expect(markdown).toContain(
			'| `dense-values` | `ascend-readxlsx-raw-values-operation-bytes` | 3.166 | 3.287 | 91.1 MiB | no optimization target |',
		)
		expect(markdown).toContain(
			'| `feature-rich` | `ascend` | 4.955 | 5.161 | 157.2 MiB | claim blocked by competitor semantic mismatch |',
		)
		expect(markdown).toContain('In a clean detached current full-profile run at `9ddfff91`')
		expect(markdown).toContain('Next action: stop production optimization from this evidence.')

		expect(markdown).toContain('## Metadata-Only Calamine Boundary')
		expect(markdown).toContain('Classification: defer/optimization target.')
		expect(markdown).toContain('Calamine is now a comparable\nmetadata-only runner')
		expect(markdown).toContain('commit `b6925afe`')
		expect(markdown).toContain(
			'/private/tmp/ascend-metadata-calamine-clean-b6925afe/metadata-calamine-head-to-head.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-metadata-calamine-clean-b6925afe/metadata-calamine-scoreboard.json',
		)
		expect(markdown).toContain(
			'| `python-calamine-metadata-only` | ran/won | 0.056 | 0.078 | 0.122 | 28.4 MiB |',
		)
		expect(markdown).toContain(
			'| `ascend-external-metadata-only` | ran/lost | 3.211 | 8.596 | 0.623 | 88.4 MiB |',
		)
		expect(markdown).toContain(
			'`profileLeaderFailures` contains the metadata-only loss to\n`python-calamine-metadata-only`',
		)
		expect(markdown).toContain('baseline median `0.105 ms`, patched median `0.168 ms`')
		expect(markdown).toContain('"Ascend beats Calamine on metadata-only open."')

		expect(markdown).toContain('## Cycle: Current Metadata-Only Calamine Recheck')
		expect(markdown).toContain('Classification: kill/defer.')
		expect(markdown).toContain('commit `c70385bc`')
		expect(markdown).toContain(
			'/private/tmp/ascend-metadata-current-c70385bc-runs/metadata-calamine-head-to-head.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-metadata-current-c70385bc-runs/metadata-calamine-scoreboard.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-metadata-current-c70385bc-patched-runs/metadata-calamine-head-to-head.json',
		)
		expect(markdown).toContain(
			'| `python-calamine-metadata-only` | ran/won | 0.070 | 0.151 | 0.309 | 28.4 MiB |',
		)
		expect(markdown).toContain(
			'| `ascend-external-metadata-only` | ran/lost vs Calamine, ran/won vs SheetJS/openpyxl | 1.208 | 6.171 | 0.840 | 83.4 MiB |',
		)
		expect(markdown).toContain('`winner=python-calamine-metadata-only expected=ascend`')
		expect(markdown).toContain(
			'patched Ascend median\n`1.348 ms`, p95 `4.410 ms`, CV `0.628`, peak RSS `90.6 MiB`',
		)
		expect(markdown).toContain(
			'Semantic boundary: this row is comparable only for the generated plain\nmetadata-only sheet-list contract.',
		)
		expect(markdown).toContain('"Calamine proves faster safe-open trust inspection than Ascend."')
		expect(markdown).toContain('kill the capsule-skip optimization target')

		expect(markdown).toContain(
			'## Cycle: Metadata-Only Relationship Recovery Profile at `38cd8ec5`',
		)
		expect(markdown).toContain('Classification: kill/defer.')
		expect(markdown).toContain(
			'Profiling named one narrow production cost\ncenter, but the smallest safe candidate did not validate',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-calamine-head-to-head.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-metadata-profile-38cd8ec5-runs/metadata-ascend-only-repeat40.json',
		)
		expect(markdown).toContain(
			'| `python-calamine-metadata-only` | ran/won | 0.102 | 1.692 | 1.705 | 28.7 MiB |',
		)
		expect(markdown).toContain(
			'| `ascend-external-metadata-only` | ran/lost vs Calamine and SheetJS, ran/won vs OpenPyXL | 3.558 | 8.432 | 0.594 | 88.8 MiB |',
		)
		expect(markdown).toContain('`recoverWorkbookRelationships`, `availablePartsForContentType`')
		expect(markdown).toContain(
			'baseline median `2.779 ms`, p95 `13.191 ms`, CV `1.046`, peak RSS\n`81.4 MiB`; patched median `3.017 ms`, p95 `13.135 ms`, CV `0.949`, peak RSS\n`83.9 MiB`',
		)
		expect(markdown).toContain(
			'"The relationship-recovery early-return patch improves metadata-only reads."',
		)

		expect(markdown).toContain('## Cycle: Dense Values Write SOTA Gate')
		expect(markdown).toContain('Classification: comparable external evidence plus defer.')
		expect(markdown).toContain('commit\n`4b8b82b6`')
		expect(markdown).toContain(
			'/private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-head-to-head.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-write-dense-current-4b8b82b6-runs/write-dense-values-fastest-repeat15.json',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/lost | 70.080 | 222.688 | 0.674 | 58.8 MiB | 172260 |',
		)
		expect(markdown).toContain(
			'| `rust-xlsxwriter` | ran/won vs Ascend | 26.806 | 31.020 | 0.076 | 17.3 MiB | 119133 |',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won | 4.851 | 9.392 | 0.244 | 73.7 MiB | 172260 |',
		)
		expect(markdown).toContain(
			'| `rust-xlsxwriter` | ran/lost vs Ascend | 31.758 | 36.138 | 0.070 | 18.8 MiB | 119134 |',
		)
		expect(markdown).toContain('`winner=rust-xlsxwriter expected=ascend`')
		expect(markdown).toContain(
			'Focused repeat-15 fastest-writer rerun: `profileLeaderFailures: []`',
		)
		expect(markdown).toContain('`closedxml` was `runner unavailable`')
		expect(markdown).toContain('"Ascend is SOTA for XLSX write."')
		expect(markdown).toContain('defer production optimization from this row')

		expect(markdown).toContain('## Cycle: Dense Values TS/JS/Rust Write Head-to-Head at `7cc7e2c3`')
		expect(markdown).toContain(
			'Classification: comparable external evidence plus defer. This refreshes the\ngenerated dense-value write row against the TS/JS and Rust libraries',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-write-js-rust-current-7cc7e2c3-runs/write-dense-values-js-rust-repeat15.json',
		)
		expect(markdown).toContain('SheetJS `0.18.5`, ExcelJS `4.4.0`, rust_xlsxwriter `0.1.0`')
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won | 6.013 | 7.776 | 0.103 | 81.0 MiB | 172259 |',
		)
		expect(markdown).toContain(
			'| `rust-xlsxwriter` | ran/lost vs Ascend | 70.629 | 148.107 | 0.343 | 21.0 MiB | 119134 |',
		)
		expect(markdown).toContain(
			'| `sheetjs` | ran/lost vs Ascend | 65.011 | 119.687 | 0.253 | 240.6 MiB | 1181431 |',
		)
		expect(markdown).toContain(
			'| `exceljs` | ran/lost vs Ascend | 124.324 | 188.872 | 0.169 | 244.2 MiB | 121315 |',
		)
		expect(markdown).toContain(
			'Full `xlsx-write-sota` coverage still fails, with 59 coverage failures',
		)
		expect(markdown).toContain('rust_xlsxwriter and ExcelJS emit smaller\nXLSX files')
		expect(markdown).toContain('"Ascend beats every TS/JS or Rust writer on every workload."')
		expect(markdown).toContain(
			'"Ascend uses less memory than rust_xlsxwriter on dense-value writes."',
		)

		expect(markdown).toContain('## Cycle: Sparse Wide Write Current Tail Boundary at `6595d42c`')
		expect(markdown).toContain(
			'Classification: comparable external evidence plus p95/tail boundary.',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-write-values-current-6595d42c-runs/write-values-sparse-repeat15-p95-scoreboard.json',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | median winner, p95 loss vs `rust-xlsxwriter` | 27.860 | 173.863 | 0.919 | 117.1 MiB | 228209 |',
		)
		expect(markdown).toContain(
			'| `rust-xlsxwriter` | median loss vs Ascend, p95 winner | 41.707 | 54.644 | 0.203 | 53.6 MiB | 175581 |',
		)
		expect(markdown).toContain('P95 scoreboard: sparse-wide group winner was `rust-xlsxwriter`.')
		expect(markdown).toContain('"Ascend has the best sparse-wide write tail latency."')
		expect(markdown).toContain('Keep sparse-wide\nwrite-speed wording scoped to median')

		expect(markdown).toContain('## Cycle: Sparse Wide TS/JS/Rust Write Head-to-Head at `a328b573`')
		expect(markdown).toContain(
			'Classification: comparable external evidence plus stale-boundary update.',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-write-sparse-js-rust-current-a328b573-runs/write-sparse-wide-js-rust-repeat15-p95-scoreboard.json',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won median and p95 | 15.530 | 20.085 | 0.134 | 169.9 MiB | 228209 |',
		)
		expect(markdown).toContain(
			'| `rust-xlsxwriter` | ran/lost vs Ascend | 21.016 | 23.789 | 0.048 | 52.7 MiB | 175581 |',
		)
		expect(markdown).toContain(
			'| `sheetjs` | ran/lost vs Ascend | 402.837 | 593.206 | 0.185 | 351.3 MiB | 883673 |',
		)
		expect(markdown).toContain(
			'| `exceljs` | ran/lost vs Ascend | 3892.681 | 7766.345 | 0.353 | 1563.6 MiB | 184376 |',
		)
		expect(markdown).toContain(
			'P95 scoreboard: sparse-wide group winner was `ascend-external-writer`',
		)
		expect(markdown).toContain(
			'Full `xlsx-write-sota` coverage still fails, with 59 coverage failures',
		)
		expect(markdown).toContain(
			'"Ascend uses less memory than rust_xlsxwriter on sparse-wide writes."',
		)

		expect(markdown).toContain('## Cycle: Plain Text Write SOTA Gate')
		expect(markdown).toContain('Classification: comparable external evidence plus defer.')
		expect(markdown).toContain('commit\n`98752c84`')
		expect(markdown).toContain(
			'/private/tmp/ascend-write-plain-text-current-98752c84-runs/write-plain-text-head-to-head.json',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won | 16.937 | 29.104 | 0.313 | 64.4 MiB | 169099 |',
		)
		expect(markdown).toContain(
			'| `excelize` | ran/lost vs Ascend | 28.623 | 39.311 | 0.197 | 21.2 MiB | 142890 |',
		)
		expect(markdown).toContain(
			'| `rust-xlsxwriter` | ran/lost vs Ascend | 39.684 | 40.941 | 0.066 | 24.1 MiB | 229138 |',
		)
		expect(markdown).toContain('`profileLeaderFailures: []`')
		expect(markdown).toContain('`closedxml` failed with `CSSM_ModuleLoad()`')
		expect(markdown).toContain('"Ascend beats ClosedXML on plain-text writes."')
		expect(markdown).toContain(
			'defer production optimization from this row and continue only with\nanother existing `xlsx-write-sota` row',
		)

		expect(markdown).toContain('## Cycle: Plain Text TS/JS/Rust Write Head-to-Head at `1bd995e5`')
		expect(markdown).toContain('Classification: comparable external evidence plus defer.')
		expect(markdown).toContain(
			'/private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-write-plain-js-rust-current-1bd995e5-runs/write-plain-text-js-rust-repeat15-p95-scoreboard.json',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won median and p95 | 3.898 | 4.080 | 0.033 | 96.9 MiB | 169097 |',
		)
		expect(markdown).toContain(
			'| `sheetjs` | ran/lost vs Ascend | 29.548 | 42.209 | 0.124 | 278.4 MiB | 1832541 |',
		)
		expect(markdown).toContain(
			'| `rust-xlsxwriter` | ran/lost vs Ascend | 29.865 | 33.282 | 0.038 | 28.7 MiB | 229139 |',
		)
		expect(markdown).toContain(
			'| `exceljs` | ran/lost vs Ascend | 91.269 | 112.351 | 0.067 | 302.8 MiB | 232106 |',
		)
		expect(markdown).toContain(
			'P95 scoreboard: plain-text group winner was `ascend-external-writer`',
		)
		expect(markdown).toContain(
			'"Ascend uses less memory than rust_xlsxwriter on plain-text writes."',
		)

		expect(markdown).toContain('## Cycle: Plain Text ClosedXML/NPOI Write Coverage at `e0c41fe5`')
		expect(markdown).toContain('Classification: comparable external evidence plus blocker update.')
		expect(markdown).toContain(
			'/private/tmp/ascend-write-closedxml-current-e0c41fe5-runs/write-plain-text-closedxml-npoi-repeat15.json',
		)
		expect(markdown).toContain('`ACCEPT_NPOI_OSMF_LICENSE=1`')
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won | 4.008 | 4.203 | 0.029 | 98.3 MiB | 169097 |',
		)
		expect(markdown).toContain(
			'| `closedxml` | ran/lost vs Ascend | 102.203 | 189.116 | 0.323 | 167.5 MiB | 224794 |',
		)
		expect(markdown).toContain(
			'| `npoi` | ran/lost vs Ascend | 201.248 | 312.078 | 0.244 | 150.4 MiB | 224417 |',
		)
		expect(markdown).toContain('Full `xlsx-write-sota` coverage still fails')
		expect(markdown).toContain('"ClosedXML and NPOI broad write coverage is complete."')
		expect(markdown).toContain(
			'use current ClosedXML/NPOI runner availability to split additional\nwrite-values blockers',
		)

		expect(markdown).toContain('## Cycle: String Heavy Write SOTA Gate')
		expect(markdown).toContain('commit\n`67b900ed`')
		expect(markdown).toContain(
			'/private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-head-to-head.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-write-string-heavy-current-67b900ed-runs/write-string-heavy-fastest-repeat15.json',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/lost | 152.628 | 209.349 | 0.566 | 60.7 MiB | 201985 |',
		)
		expect(markdown).toContain(
			'| `rust-xlsxwriter` | ran/won vs Ascend | 65.802 | 135.550 | 0.424 | 25.5 MiB | 237837 |',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won | 63.896 | 372.935 | 0.927 | 65.5 MiB | 201985 |',
		)
		expect(markdown).toContain(
			'| `sheetjs` | ran/lost vs Ascend | 94.161 | 209.940 | 0.352 | 276.5 MiB | 2016032 |',
		)
		expect(markdown).toContain('Full external row: group winner was `rust-xlsxwriter`')
		expect(markdown).toContain(
			'Focused repeat-15 fastest-writer rerun: group winner was\n  `ascend-external-writer`',
		)
		expect(markdown).toContain('"Ascend has the best tail latency for string-heavy writes."')
		expect(markdown).toContain(
			'defer production optimization from this row. If string-heavy\nmatters for a release claim later',
		)

		expect(markdown).toContain('## Cycle: String Heavy Write Optimization')
		expect(markdown).toContain('Classification: validated optimization.')
		expect(markdown).toContain('Commit: `bd1629373a9a4a17edd2db9125b6cd19e6df7504`')
		expect(markdown).toContain(
			'`packages/io-xlsx/src/writer/dense-rows.ts` now uses the synchronous dense\n  ZIP builder when the estimated dense sheet XML is at or below 4 MiB.',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-write-string-heavy-optimized-bd162937-runs/write-string-heavy-fastest-repeat15.json',
		)
		expect(markdown).toContain(
			'| `string-heavy` dense streaming | 107.176 | 12.418 | samples up to 198.748 ms | samples up to 31.683 ms |',
		)
		expect(markdown).toContain(
			'| `dense-values` dense streaming | 76.335 | 11.179 | samples up to 176.338 ms | samples up to 24.411 ms |',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won | 5.001 | 6.077 | 0.099 | 91.8 MiB | 201984 |',
		)
		expect(markdown).toContain(
			'| `rust-xlsxwriter` | ran/lost vs Ascend | 30.237 | 34.996 | 0.079 | 24.3 MiB | 237837 |',
		)
		expect(markdown).toContain('Ascend now wins median and p95 on this focused row.')
		expect(markdown).toContain('"Ascend beats ClosedXML or NPOI on string-heavy writes."')
		expect(markdown).toContain(
			'continue optimizing or bounding the next existing `xlsx-write-sota`\ngap',
		)

		expect(markdown).toContain('## Cycle: String Heavy ClosedXML/NPOI Write Coverage at `6cc5076f`')
		expect(markdown).toContain('Classification: comparable external evidence plus blocker update.')
		expect(markdown).toContain(
			'/private/tmp/ascend-write-dotnet-current-6cc5076f-runs/write-string-heavy-closedxml-npoi-repeat15.json',
		)
		expect(markdown).toContain('ClosedXML `0.105.0.0`; NPOI `2.8.0.0`')
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won | 4.131 | 4.820 | 0.064 | 91.9 MiB | 201984 |',
		)
		expect(markdown).toContain(
			'| `closedxml` | ran/lost vs Ascend | 113.223 | 317.533 | 0.502 | 140.0 MiB | 246941 |',
		)
		expect(markdown).toContain(
			'| `npoi` | ran/lost vs Ascend | 206.278 | 480.695 | 0.376 | 146.6 MiB | 234032 |',
		)
		expect(markdown).toContain(
			'Full `xlsx-write-sota` coverage still fails, with 61 coverage failures',
		)
		expect(markdown).toContain(
			'ClosedXML and NPOI matched sorted semantic values\nbut not ordered semantic hashes',
		)
		expect(markdown).toContain('"Ascend beats ClosedXML or NPOI on every write workload."')
		expect(markdown).toContain('"ClosedXML and NPOI broad write coverage is complete."')

		expect(markdown).toContain(
			'## Cycle: Plain Text Workbook Writer Metadata-Key Optimization at `fd616906`',
		)
		expect(markdown).toContain(
			'Classification: validated optimization. The workbook-buffered sheet writer was',
		)
		expect(markdown).toContain(
			'`packages/io-xlsx/src/writer/sheet.ts` now computes `formulaStorageKey(row,\n  col)` only when `storedFormulaText` or `preservedCellMetadata` is non-empty.',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-baseline-repeat40.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-writer-key-current-fd616906-runs/plain-text-patched-repeat40.json',
		)
		expect(markdown).toContain(
			'| Baseline | 4.780 | 6.607 | 0.211 | 8.126 | 10.175 | 0.135 | 8.367 M cells/s | 161.8 MiB | 176828 |',
		)
		expect(markdown).toContain(
			'| Patched | 4.094 | 4.762 | 0.076 | 7.024 | 7.999 | 0.058 | 9.771 M cells/s | 161.1 MiB | 176828 |',
		)
		expect(markdown).toContain('patched write median improved by `14.36%`')
		expect(markdown).toContain('write p95 improved by\n`27.92%`')
		expect(markdown).toContain(
			'(pass) writeXlsx > preserves original stored formula text when unrelated edits dirty the sheet',
		)
		expect(markdown).toContain(
			'| `ascend-external-writer` | ran/won | 4.319 | 4.615 | 0.033 | 97.5 MiB | 169097 |',
		)
		expect(markdown).toContain(
			'| `fastexcel-java` | ran/lost vs Ascend | 46.587 | 239.669 | 0.851 | 1584.0 MiB | 227254 |',
		)
		expect(markdown).toContain('"Ascend improved every write workload."')
		expect(markdown).toContain('do not optimize this same plain-text workbook-buffered row again')

		expect(markdown).toContain('## Full Current-Commit Gate: XLSX Read SOTA')
		expect(markdown).toContain(
			'Classification: blocked/defer. No production optimization is justified',
		)
		expect(markdown).toContain('full `xlsx-read-sota` open/inspect coverage')
		expect(markdown).toContain('Commit: `4b1c1734b95ee96da0078b60f5e439768303e04e`')
		expect(markdown).toContain('`leaderFailures: []`')
		expect(markdown).toContain('`profileLeaderFailures: []`')
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-4b1c1734-runs/xlsx-read-sota-all.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-4b1c1734-runs/xlsx-read-sota-all-scoreboard.json',
		)
		expect(markdown).toContain(
			'ClosedXML` missing for `dense-values`, `sparse-wide`, `string-heavy`',
		)
		expect(markdown).toContain('`SheetJS` and `Calamine` ineligible for `feature-rich`')
		expect(markdown).toContain(
			'`metadata-only` reports `missing-comparable` for required competitors',
		)
		expect(markdown).toContain(
			'`selected-sheet` is `unsupported-operation` for\n  `ExcelJS`, `Apache POI`, and `ClosedXML`',
		)
		expect(markdown).toContain('| Completed comparable profile rows | ran/won |')
		expect(markdown).toContain('| ClosedXML | blocked |')
		expect(markdown).toContain('| fastxlsx | runner unavailable |')
		expect(markdown).toContain(
			'"Ascend beats ClosedXML, fastxlsx, unsupported runners, or semantic-mismatch rows."',
		)

		expect(markdown).toContain('## Cycle: ClosedXML Focused Head-to-Head Read')
		expect(markdown).toContain('ClosedXML is no longer a runner blocker')
		expect(markdown).toContain('Commit: `2f5c17617ae2c41cac84558edebe3b3174c30a09`')
		expect(markdown).toContain('ClosedXML runner version: `0.105.0.0`')
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-2f5c1761-runs/ascend-closedxml-head-to-head-all.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-2f5c1761-runs/ascend-closedxml-head-to-head-all-scoreboard.json',
		)
		expect(markdown).toContain(
			'| `dense-values` | `ascend-readxlsx-raw-values-operation-bytes` | 9.394 ms / 11.905 ms / 0.165 / 85.6 MiB | 369.344 ms / 884.345 ms / 0.514 / 113.4 MiB | ran/won |',
		)
		expect(markdown).toContain(
			'| `feature-rich` | `ascend-readxlsx-values-rich-metadata-bytes` | 50.600 ms / 52.510 ms / 0.036 / 169.3 MiB | 194.255 ms / 610.081 ms / 0.618 / 142.5 MiB | ran/won |',
		)
		expect(markdown).toContain(
			'| `selected-sheet` | n/a | n/a | n/a | not comparable | ClosedXML remains `unsupported-operation`',
		)
		expect(markdown).toContain(
			'| `metadata-only` | n/a | n/a | n/a | not comparable | ClosedXML remains `unsupported-operation`',
		)
		expect(markdown).toContain('`leaderFailures: []`')
		expect(markdown).toContain('`profileLeaderFailures: []`')
		expect(markdown).toContain(
			'"Ascend beats ClosedXML for selected-sheet or metadata-only reads."',
		)
		expect(markdown).toContain('no\nlonger ClosedXML value-read coverage')

		expect(markdown).toContain('## Cycle: Selected-Sheet OpenPyXL Head-to-Head Read')
		expect(markdown).toContain(
			'The openpyxl selected-sheet gap moved from\n`unsupported-operation` to a passing measured row',
		)
		expect(markdown).toContain('Commit: `57c5a2420d9d19a3f4f2138bf6644303450b01a1`')
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-57c5a242-runs/selected-sheet-openpyxl-head-to-head.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-57c5a242-runs/selected-sheet-openpyxl-head-to-head-scoreboard.json',
		)
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend` | 37.985 | 82.642 | 0.726 | 185.6 MiB |',
		)
		expect(markdown).toContain('| SheetJS | ran/lost | `sheetjs` | 66.446')
		expect(markdown).toContain('| openpyxl | ran/lost | `openpyxl` | 372.944')
		expect(markdown).toContain(
			'`requiredCompetitors=Ascend,SheetJS,openpyxl` because the current evidence',
		)
		expect(markdown).toContain(
			'openpyxl selected-sheet `unsupported-operation` coverage gap is removed',
		)
		expect(markdown).toContain('"Ascend has a full selected-sheet SOTA claim."')
		expect(markdown).toContain('Add external-process selected-sheet\nlanes for Ascend and SheetJS')

		expect(markdown).toContain('## Cycle: Selected-Sheet Same-Lane External Read')
		expect(markdown).toContain('mixed in-process/external timing lane')
		expect(markdown).toContain('Commit: `3916386295e83b234dba10bdb7f007a9f5d52704`')
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-39163862-runs/selected-sheet-same-lane.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-39163862-runs/selected-sheet-same-lane-scoreboard.json',
		)
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend-external-values` | 22.066 | 25.622 | 0.085 | 154.4 MiB |',
		)
		expect(markdown).toContain('| SheetJS | ran/lost | `sheetjs` | 27.981')
		expect(markdown).toContain('| openpyxl | ran/lost | `openpyxl` | 205.474')
		expect(markdown).toContain(
			'No selected-sheet `coverageFailures` remain for Ascend, SheetJS, or openpyxl.',
		)
		expect(markdown).toContain(
			'Selected-sheet `coverageGaps` remain for ExcelJS, Apache POI, and ClosedXML',
		)
		expect(markdown).toContain(
			'"Ascend has a full selected-sheet SOTA claim across every library."',
		)
		expect(markdown).toContain('metadata-only same-lane coverage')

		expect(markdown).toContain(
			'bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all',
		)
		expect(markdown).toContain(
			'bun run fixtures/benchmarks/competitive-scoreboard.ts <suite.json> --json --metric medianMs --require-profile xlsx-read-sota',
		)
		expect(markdown).toContain(
			'bun run fixtures/benchmarks/competitive-scoreboard.ts /private/tmp/ascend-external-claim-matrix-2026-05-15-clean/string-heavy-read-values-clean.json --json --metric medianMs --require-profile xlsx-read-sota',
		)

		for (const workload of REQUIRED_WORKLOAD_GAPS) expect(markdown).toContain(workload)
		for (const runner of UNAVAILABLE_RUNNER_ROWS) {
			expect(markdown).toContain(`| ${runner}`)
			expect(markdown).toContain('runner unavailable')
		}
		expect(markdown).toContain('| ClosedXML | blocked |')

		expect(markdown).toContain('repeat 5')
		expect(markdown).toContain('1 warmup')
		expect(markdown).toContain('median 9.347 ms')
		expect(markdown).toContain('p95 16.545 ms')
		expect(markdown).toContain('CV 0.300')
		expect(markdown).toContain('129.8 MiB')
		expect(markdown).toContain('Darwin 25.4.0 arm64')
		expect(markdown).toContain('SheetJS `xlsx@0.18.5`')
		expect(markdown).toContain('ExcelJS `4.4.0`')

		expect(markdown).toContain('## Cycle: Sparse-Wide Value Read')
		expect(markdown).toContain('Classification: defer. No production optimization is justified')
		expect(markdown).toContain('generated `sparse-wide` workbook')
		expect(markdown).toContain('Commit: `2e71900f`')
		expect(markdown).toContain('5000 rows x 256 columns')
		expect(markdown).toContain('23,093 populated cells')
		expect(markdown).toContain('Median ms | P95 ms | CV')
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 11.651 | 20.132 | 0.344 | 120.9 MiB |',
		)
		expect(markdown).toContain('| FastExcel Java | blocked | `fastexcel-java` | n/a')
		expect(markdown).toContain('| ClosedXML | blocked | `closedxml` | n/a')
		expect(markdown).toContain('| fastxlsx | runner unavailable | `fastxlsx` | n/a')
		expect(markdown).toContain(
			'"Ascend beats FastExcel Java, ClosedXML, or fastxlsx" from this run.',
		)
		expect(markdown).toContain('Continue profile expansion with `styles-heavy`')

		expect(markdown).toContain('## Cycle: Styles-Heavy Value Read')
		expect(markdown).toContain('Classification: defer. No production optimization is justified')
		expect(markdown).toContain('generated `styles-heavy` workbook')
		expect(markdown).toContain('Commit: `b1a53ea0`')
		expect(markdown).toContain('2000 rows x 20 columns')
		expect(markdown).toContain('257,177 input bytes')
		expect(markdown).toContain('Median ms | P95 ms | CV')
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 5.733 | 7.387 | 0.142 | 99.4 MiB |',
		)
		expect(markdown).toContain('| ClosedXML | blocked | `closedxml` | n/a')
		expect(markdown).toContain('| fastxlsx | runner unavailable | `fastxlsx` | n/a')
		expect(markdown).toContain('"Ascend beats ClosedXML or fastxlsx" from this run.')
		expect(markdown).toContain('Continue profile expansion with `formula-heavy`')

		expect(markdown).toContain('## Cycle: Formula-Heavy Value Read')
		expect(markdown).toContain('generated `formula-heavy` workbook')
		expect(markdown).toContain('Commit: `e1c69a32`')
		expect(markdown).toContain('246,241 input bytes')
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 5.587 | 5.689 | 0.013 | 101.9 MiB |',
		)
		expect(markdown).toContain(
			'"Ascend proves formula calculation, recalc, or formula preservation" from this value-read run.',
		)
		expect(markdown).toContain('Continue profile expansion with `table-heavy`')

		expect(markdown).toContain('## Cycle: Table-Heavy Value Read')
		expect(markdown).toContain('generated `table-heavy` workbook')
		expect(markdown).toContain('Commit: `65520519`')
		expect(markdown).toContain('151,994 input bytes')
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend-readxlsx-raw-values-operation-path` | 8.822 | 9.551 | 0.052 | 115.2 MiB |',
		)
		expect(markdown).toContain('| Polars calamine | not comparable | `polars-calamine`')
		expect(markdown).toContain('| fastexcel Python | not comparable | `fastexcel`')
		expect(markdown).toContain(
			'"Ascend beats fastexcel Python, Polars calamine, Polars xlsx2csv, or Polars openpyxl" from this run.',
		)
		expect(markdown).toContain('Continue profile expansion with `feature-rich`')

		expect(markdown).toContain('## Cycle: Feature-Rich Metadata Read')
		expect(markdown).toContain('Classification: optimize, then validated ran/won')
		expect(markdown).toContain('generated `feature-rich` workbook')
		expect(markdown).toContain('Commit: `05656d4e`')
		expect(markdown).toContain('114,404 input bytes')
		expect(markdown).toContain(
			'pre-optimization clean run at detached commit `222c4898` made the comparable Ascend rich-metadata row eligible but slower than Apache POI',
		)
		expect(markdown).toContain('`ascend-readxlsx-values-rich-metadata-bytes` median 92.937 ms')
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend-readxlsx-values-rich-metadata-bytes` | 41.335 | 41.442 | 0.048 | 171.3 MiB |',
		)
		expect(markdown).toContain('| Apache POI | ran/lost | `apache-poi` | 86.697')
		expect(markdown).toContain('| SheetJS | not comparable | `sheetjs` | 33.529')
		expect(markdown).toContain(
			'"Ascend beats SheetJS, fastexcel, Calamine, FastExcel Java, Polars, or pyopenxlsx" from this rich-metadata run.',
		)
		expect(markdown).toContain('Continue profile expansion with `selected-sheet`')

		expect(markdown).toContain('## Cycle: SheetJS Feature-Rich Rich Metadata')
		expect(markdown).toContain('enable the public `bookFiles` option')
		expect(markdown).toContain('Commit: `15119c8d828e52493866c760c7fe28972e4a3bee`')
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-sheetjs-inprocess.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-sheetjs-inprocess-scoreboard.json',
		)
		expect(markdown).toContain(
			'Both rows assert 1 comment, 1 hyperlink,\n1 data validation, 1 conditional format, 1 defined name',
		)
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend` | 5.123 | 5.575 | 0.048 | 222.7 MiB |',
		)
		expect(markdown).toContain(
			'| SheetJS | ran/lost | `sheetjs` | 31.211 | 32.814 | 0.035 | 328.1 MiB |',
		)
		expect(markdown).toContain('| Calamine-family runners | not comparable | n/a')
		expect(markdown).toContain('The `feature-rich` group winner was `ascend`.')
		expect(markdown).toContain(
			'"Ascend beats Calamine-family readers on feature-rich rich-metadata reads."',
		)

		expect(markdown).toContain('## Cycle: Calamine Feature-Rich Rich Metadata Boundary')
		expect(markdown).toContain('Classification: not comparable/defer')
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-calamine-boundary.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-15119c8d-runs/feature-rich-calamine-boundary-scoreboard.json',
		)
		expect(markdown).toContain(
			'| python-calamine | not comparable | `python-calamine` | 24.332 | 39.131 | 0.246 | 42.4 MiB |',
		)
		expect(markdown).toContain(
			'| rust-calamine | not comparable | `rust-calamine` | 135.078 | 162.012 | 0.280 | 15.6 MiB |',
		)
		expect(markdown).toContain(
			'The Calamine timing lane has no winner because both rows are\n  `semantic-mismatch`.',
		)
		expect(markdown).toContain('"Calamine lost the feature-rich rich-metadata benchmark."')
		expect(markdown).toContain('kill Calamine-family rich-metadata speed comparisons')

		expect(markdown).toContain('## Cycle: Selected-Sheet Value Read')
		expect(markdown).toContain('generated `selected-sheet` workbook')
		expect(markdown).toContain('Commit: `5055d794`')
		expect(markdown).toContain('114,747 input bytes')
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend` | 3.054 | 3.156 | 0.042 | 198.0 MiB |',
		)
		expect(markdown).toContain('| SheetJS | ran/lost | `sheetjs` | 29.149')
		expect(markdown).toContain('| ExcelJS | not comparable | n/a')
		expect(markdown).toContain(
			'"Ascend beats ExcelJS, openpyxl, Calamine, Apache POI, or ClosedXML" from this selected-sheet run.',
		)
		expect(markdown).toContain('the OpenPyXL unsupported-operation status above is historical')
		expect(markdown).toContain('the Calamine unsupported-operation status above is also')
		expect(markdown).toContain('Commit `79d6cefd` proves the\npython-calamine runner')
		expect(markdown).toContain('## Fold-In: OpenPyXL Selected-Sheet Projection')
		expect(markdown).toContain('Classification: accepted evidence plus benchmark blocker.')
		expect(markdown).toContain('OpenPyXL reported\n  median `204.707 ms`')
		expect(markdown).toContain(
			'"Ascend beats OpenPyXL for selected-sheet reads" from the current worktree',
		)
		expect(markdown).toContain(
			'bun run fixtures/benchmarks/competitive-io.ts --json --category read --competitor all --libraries ascend,sheetjs,openpyxl --workload selected-sheet',
		)
		expect(markdown).toContain(
			'keeps Ascend-vs-OpenPyXL selected-sheet\nspeed wording forbidden because the timing lanes differ',
		)
		expect(markdown).toContain('Continue profile expansion with `metadata-only`')

		expect(markdown).toContain('## Cycle: Metadata-Only Read')
		expect(markdown).toContain('generated `metadata-only` workbook')
		expect(markdown).toContain('Commit: `5261b08d`')
		expect(markdown).toContain('15,347 input bytes')
		expect(markdown).toContain(
			'| Ascend external | ran/won | `ascend-external-metadata-only-bytes` | 0.310 | 0.329 | 0.082 | 89.8 MiB |',
		)
		expect(markdown).toContain('| openpyxl | ran/lost | `openpyxl-metadata-only` | 1.986')
		expect(markdown).toContain('| SheetJS | ran/lost | `sheetjs` | 0.832')
		expect(markdown).toContain(
			'"Ascend beats ExcelJS, Calamine, Apache POI, or ClosedXML" from this metadata-only run.',
		)
		expect(markdown).toContain('Continue profile expansion with `warm-workflow`')

		expect(markdown).toContain('## Cycle: Metadata-Only Same-Lane External Read')
		expect(markdown).toContain('mixed a preloaded-bytes\nAscend external lane')
		expect(markdown).toContain('Commit: `fa3a13dc1f72de489d6d301bf1f81cbe3400df0f`')
		expect(markdown).toContain('fixtures/benchmarks/runners/metadata-only-readers.manifest.json')
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-fa3a13dc-runs/metadata-only-same-lane.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-fa3a13dc-runs/metadata-only-same-lane-scoreboard.json',
		)
		expect(markdown).toContain(
			'| Ascend | ran/won | `ascend-external-metadata-only` | 0.394 | 0.457 | 0.148 | 91.2 MiB |',
		)
		expect(markdown).toContain('| SheetJS | ran/lost | `sheetjs-metadata-only` | 0.941')
		expect(markdown).toContain('| openpyxl | ran/lost | `openpyxl-metadata-only` | 2.050')
		expect(markdown).toContain(
			'No metadata-only `coverageFailures` remain for Ascend, SheetJS, or openpyxl.',
		)
		expect(markdown).toContain(
			'Metadata-only `coverageGaps` remain for ExcelJS, Calamine, Apache POI, and\n  ClosedXML',
		)
		expect(markdown).toContain('"Ascend has a full metadata-only SOTA claim across every library."')
		expect(markdown).toContain('feature-rich semantic mismatches for SheetJS and Calamine')

		expect(markdown).toContain('## Cycle: FastXLSX Cell-Materialization Coverage')
		expect(markdown).toContain('historical FastXLSX row was a cooked non-result')
		expect(markdown).toContain('Commit: `52f7f172e2826211a4a0ba811a722077dcd1a824`')
		expect(markdown).toContain('fastxlsx==0.2.0')
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-52f7f172-runs/fastxlsx-cell-materialization-all.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-perf-hillclimb-52f7f172-runs/fastxlsx-cell-materialization-all-scoreboard.json',
		)
		expect(markdown).toContain(
			'| `dense-values` | ran/won | 9.270 ms / 10.247 ms / 0.067 / 102.6 | ran/lost | 34.884 ms / 57.856 ms / 0.287 / 52.5 |',
		)
		expect(markdown).toContain(
			'| `sparse-wide` | ran/won | 13.066 ms / 14.376 ms / 0.047 / 127.3 | ran/lost | 101.486 ms / 102.578 ms / 0.009 / 87.5 |',
		)
		expect(markdown).toContain(
			'| `feature-rich` | not comparable | 11.573 ms / 11.793 ms / 0.017 / 116.6 | not comparable | 25.734 ms / 25.911 ms / 0.006 / 53.3 |',
		)
		expect(markdown).toContain('Every comparable FastXLSX value/warm group winner was')
		expect(markdown).toContain('"Ascend beats FastXLSX on rich-metadata feature-rich reads."')
		expect(markdown).toContain('Carry the isolated Python 3.12\nFastXLSX setup')

		expect(markdown).toContain('## Cycle: Current FastXLSX Carry-Forward Gate')
		expect(markdown).toContain('Classification: comparable external evidence plus defer')
		expect(markdown).toContain('current commit `248f76d9`')
		expect(markdown).toContain(
			'/private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-cell-materialization-all.json',
		)
		expect(markdown).toContain(
			'/private/tmp/ascend-fastxlsx-current-248f76d9/fastxlsx-table-heavy-repeat15.json',
		)
		expect(markdown).toContain(
			'| `dense-values` | ran/won | 19.885 ms / 25.612 ms / 0.164 / 101.3 MiB | ran/lost | 42.789 ms / 51.235 ms / 0.092 / 56.6 MiB |',
		)
		expect(markdown).toContain(
			'| `feature-rich` | not comparable | 25.415 ms / 29.076 ms / 0.093 / 114.9 MiB | not comparable | 58.247 ms / 67.872 ms / 0.084 / 53.2 MiB |',
		)
		expect(markdown).toContain(
			'| `ascend-readxlsx-cell-materialization-bytes` | ran/won | 30.133 | 38.316 | 0.152 | 140.5 MiB |',
		)
		expect(markdown).toContain('| `fastxlsx` | ran/lost | 46.843 | 59.085 | 0.106 | 55.6 MiB |')
		expect(markdown).toContain(
			'Focused all-workload `--assert-leader ascend`: `leaderFailures: []`',
		)
		expect(markdown).toContain(
			'Table-heavy repeat-15 `--assert-leader ascend`: `leaderFailures: []`',
		)
		expect(markdown).toContain(
			'The all-workload table-heavy repeat-5 group briefly had `fastxlsx` as the\n  median winner by 1.7%',
		)
		expect(markdown).toContain('FastXLSX consistently used lower RSS than Ascend on this lane.')
		expect(markdown).toContain('"Ascend beats FastXLSX on memory."')
		expect(markdown).toContain(
			'Next action: defer production optimization on FastXLSX value materialization.',
		)

		expect(markdown).toContain(
			'## Cycle: Tracked Real Workbook Strings/Links Open Boundary at `e8654a0b`',
		)
		expect(markdown).toContain(
			'Classification: scoped real-workbook evidence plus claim downgrade.',
		)
		expect(markdown).toContain('fixtures/xlsx/xlsxwriter/strings_links.xlsx')
		expect(markdown).toContain('e46b7e597607b4d4819ae83265f8d160904e7b01537637db68bff698c46d522b')
		expect(markdown).toContain(
			'/private/tmp/ascend-real-workbook-current-e8654a0b-runs/xlsxwriter-strings-links-read-nonordered-repeat15.json',
		)
		expect(markdown).toContain('Used range: `Strings!A2:D200`')
		expect(markdown).toContain(
			'| `rust-calamine` | ran/won on its lane | 0.650 | 0.759 | 0.073 | 3.0 MiB |',
		)
		expect(markdown).toContain(
			'| `ascend-external-values` | ran/won on its lane | 2.162 | 2.813 | 0.109 | 114.7 MiB |',
		)
		expect(markdown).toContain(
			'| `openpyxl-read-only-values` | ran/won on its lane, slower than Ascend | 5.367 | 5.727 | 0.040 | 48.9 MiB |',
		)
		expect(markdown).toContain('`external-internal-file-path-materialization-timing`')
		expect(markdown).toContain('Rust Calamine is faster on the narrower materialization lane')
		expect(markdown).toContain('"Ascend is fastest for real-workbook open/inspect."')
		expect(markdown).toContain('do not collapse them\ninto a single cross-library leaderboard')

		expect(markdown).toContain('## Cycle: Warm Workflow Value Read')
		expect(markdown).toContain('generated `warm-workflow` workbook')
		expect(markdown).toContain('Commit: `add13c79`')
		expect(markdown).toContain('112,699 input bytes')
		expect(markdown).toContain(
			'| Ascend operation bytes | ran/won | `ascend-readxlsx-raw-values-operation-bytes` | 3.244 | 3.297 | 0.016 | 92.5 MiB |',
		)
		expect(markdown).toContain('| FastExcel Java | ran/lost | `fastexcel-java` | 14.338')
		expect(markdown).toContain('| ClosedXML | blocked | `closedxml` | n/a')
		expect(markdown).toContain('| fastxlsx | runner unavailable | `fastxlsx` | n/a')
		expect(markdown).toContain('Run or assemble a current-commit full-profile gate next')

		expect(markdown).toContain('Promote: no.')
		expect(markdown).toContain(
			'Optimize: no production optimization from the partial profile rows.',
		)
		expect(markdown).toContain('Defer: yes.')
		expect(markdown).toContain(
			'The immediate next action is a current-commit full-profile gate or merged profile artifact, plus runner hardening for ClosedXML, fastxlsx, rich-metadata semantic mismatches, selected-sheet unsupported-operation gaps, and metadata-only unsupported-operation gaps',
		)
	})
})

function readRepoFile(path: string): string {
	return readFileSync(new URL(path, REPO_ROOT), 'utf-8')
}
