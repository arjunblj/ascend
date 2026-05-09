import type { BenchmarkSuiteResult } from './results.ts'

export interface TuiPerformanceTarget {
	readonly scenario: string
	readonly metric: 'medianMs' | 'p95Ms'
	readonly max: number
}

export const tuiPerformanceTargets: readonly TuiPerformanceTarget[] = [
	{ scenario: 'file-hub-first-paint', metric: 'p95Ms', max: 16 },
	{ scenario: 'warm-grid-navigation', metric: 'p95Ms', max: 16 },
	{ scenario: 'formula-entry-commit', metric: 'p95Ms', max: 16 },
	{ scenario: 'formula-edit-cursor-f4', metric: 'p95Ms', max: 16 },
	{ scenario: 'formula-trace-workflow', metric: 'p95Ms', max: 16 },
	{ scenario: 'formula-point-mode-workflow', metric: 'p95Ms', max: 16 },
	{ scenario: 'paste-10k-cells', metric: 'p95Ms', max: 250 },
	{ scenario: 'command-palette-search', metric: 'p95Ms', max: 16 },
	{ scenario: 'terminal-calibration-workflow', metric: 'p95Ms', max: 16 },
	{ scenario: 'terminal-input-stream-parser', metric: 'medianMs', max: 16 },
	{ scenario: 'dialog-command-workflow', metric: 'p95Ms', max: 16 },
	{ scenario: 'table-comment-workflow', metric: 'p95Ms', max: 16 },
	{ scenario: 'object-dialog-workflow', metric: 'p95Ms', max: 16 },
	{ scenario: 'object-inspector-workflow', metric: 'p95Ms', max: 16 },
	{ scenario: 'print-preview-workflow', metric: 'p95Ms', max: 16 },
	{ scenario: 'file-save-export-workflow', metric: 'p95Ms', max: 100 },
	{ scenario: 'find-replace-workflow', metric: 'medianMs', max: 16 },
	{ scenario: 'dialog-form-apply', metric: 'p95Ms', max: 16 },
	{ scenario: 'resize-frame', metric: 'p95Ms', max: 20 },
	{ scenario: 'renderer-bakeoff-ansi-baseline', metric: 'p95Ms', max: 16 },
	{ scenario: 'renderer-bakeoff-opentui-line-adapter', metric: 'medianMs', max: 50 },
	{ scenario: 'metadata-grid-paint-1m-x-20', metric: 'medianMs', max: 300 },
]

export interface TuiTargetCheckResult {
	readonly target: TuiPerformanceTarget
	readonly actual: number | null
	readonly passed: boolean
	readonly skipped?: boolean
}

export function checkTuiTargets(
	suite: BenchmarkSuiteResult,
	options: { readonly skipMissing?: boolean } = {},
): readonly TuiTargetCheckResult[] {
	return tuiPerformanceTargets.map((target) => {
		const scenario = suite.cases.find((entry) => entry.name === target.scenario)
		const actual = scenario ? metricValue(scenario, target.metric) : null
		const skipped = actual === null && options.skipMissing === true
		return {
			target,
			actual,
			passed: skipped || (actual !== null && actual <= target.max),
			...(skipped ? { skipped } : {}),
		}
	})
}

export function formatTuiTargetResults(results: readonly TuiTargetCheckResult[]): string {
	const lines = ['TUI Performance Target Check', '='.repeat(72)]
	for (const result of results) {
		const status = result.skipped ? 'SKIP' : result.passed ? 'PASS' : 'FAIL'
		const actual = result.actual === null ? 'no data' : `${result.actual.toFixed(2)}`
		lines.push(
			`  [${status}] ${result.target.scenario.padEnd(28)} ${result.target.metric.padEnd(12)} ${actual.padStart(10)} <= ${result.target.max}`,
		)
	}
	const passed = results.filter((entry) => entry.passed && !entry.skipped).length
	const skipped = results.filter((entry) => entry.skipped).length
	lines.push('-'.repeat(72))
	lines.push(
		`  ${passed}/${results.length - skipped} targets met${skipped > 0 ? ` (${skipped} skipped)` : ''}`,
	)
	return lines.join('\n')
}

function metricValue(
	scenario: BenchmarkSuiteResult['cases'][number],
	metric: TuiPerformanceTarget['metric'],
): number | null {
	switch (metric) {
		case 'medianMs':
			return scenario.metrics.medianMs
		case 'p95Ms':
			return scenario.metrics.p95Ms
	}
}
