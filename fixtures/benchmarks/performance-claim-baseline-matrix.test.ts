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
			'No broad XLSX read, SOTA, or QSS-leapfrog speed claim is promotable',
		)
		expect(markdown).toContain(
			'ClosedXML was subsequently unblocked in a focused clean head-to-head run and is recorded as ran/lost, not as blocked.',
		)
		expect(markdown).toContain(
			'Current harness evidence now supports an OpenPyXL selected-sheet projection row.',
		)
		for (const workload of RECORDED_WORKLOADS) expect(markdown).toContain(`\`${workload}\``)
		expect(markdown).toContain('Humble allowed wording:')
		expect(markdown).toContain('Forbidden wording:')
		expect(markdown).toContain('"Ascend is the fastest XLSX reader."')
		expect(markdown).toContain('"Ascend is SOTA for XLSX read."')
		expect(markdown).toContain('Any wording that treats failed or unavailable runners as wins.')

		expect(markdown).toContain('## Owner-Ready Benchmark Blocker')
		expect(markdown).toContain('Owner: benchmarking/external baselines.')
		expect(markdown).toContain('broad read-speed and QSS-leapfrog performance wording is blocked')
		expect(markdown).toContain(
			'do not optimize further from the measured winning rows\n`dense-values`, `sparse-wide`, `styles-heavy`, `formula-heavy`, `table-heavy`,\n`selected-sheet`, `metadata-only`, `warm-workflow`, and `string-heavy`',
		)
		expect(markdown).toContain(
			'Failed, missing, or semantically mismatched runners are not counted as wins.',
		)

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
			'`selected-sheet` is `unsupported-operation` for\n  `ExcelJS`, `openpyxl`, `Calamine`, `Apache POI`, and `ClosedXML`',
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
			'Selected-sheet `coverageGaps` remain for ExcelJS, Calamine, Apache POI, and\n  ClosedXML',
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
