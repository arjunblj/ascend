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

describe('performance claim baseline matrix', () => {
	test('pins scoped read evidence as a defer decision, not a speed claim', () => {
		const markdown = readRepoFile(MATRIX_PATH)

		expect(markdown).toContain('Status: defer.')
		expect(markdown).toContain('This document is the tracked release claim artifact')
		expect(markdown).toContain(
			'No broad XLSX read, SOTA, or QSS-leapfrog speed claim is promotable',
		)
		expect(markdown).toContain('Humble allowed wording:')
		expect(markdown).toContain('Forbidden wording:')
		expect(markdown).toContain('"Ascend is the fastest XLSX reader."')
		expect(markdown).toContain('"Ascend is SOTA for XLSX read."')
		expect(markdown).toContain('Any wording that treats failed or unavailable runners as wins.')

		expect(markdown).toContain('## Owner-Ready Benchmark Blocker')
		expect(markdown).toContain('Owner: benchmarking/external baselines.')
		expect(markdown).toContain('broad read-speed and QSS-leapfrog performance wording is blocked')
		expect(markdown).toContain('do not optimize from this single `string-heavy` row')
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

		expect(markdown).toContain('Promote: no.')
		expect(markdown).toContain('Optimize: no production optimization from this single row.')
		expect(markdown).toContain('Defer: yes.')
	})
})

function readRepoFile(path: string): string {
	return readFileSync(new URL(path, REPO_ROOT), 'utf-8')
}
