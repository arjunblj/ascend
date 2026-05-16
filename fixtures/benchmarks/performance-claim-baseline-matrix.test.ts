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
			'do not optimize further from the partial winning rows\n`dense-values`, `sparse-wide`, `styles-heavy`, `formula-heavy`, `table-heavy`,\nand `string-heavy`',
		)
		expect(markdown).toContain(
			'Failed, missing, or semantically mismatched runners are not counted as wins.',
		)

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
		expect(markdown).toContain('Continue profile expansion with `metadata-only`')

		expect(markdown).toContain('Promote: no.')
		expect(markdown).toContain(
			'Optimize: no production optimization from the partial profile rows.',
		)
		expect(markdown).toContain('Defer: yes.')
		expect(markdown).toContain(
			'The immediate next action is `metadata-only` profile expansion plus runner hardening for FastExcel Java on `sparse-wide`, ClosedXML, fastxlsx, rich-metadata semantic mismatches, and selected-sheet unsupported-operation gaps',
		)
	})
})

function readRepoFile(path: string): string {
	return readFileSync(new URL(path, REPO_ROOT), 'utf-8')
}
