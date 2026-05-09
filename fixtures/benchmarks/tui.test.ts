import { describe, expect, test } from 'bun:test'
import { createBenchmarkSuite } from './results.ts'
import { checkTuiTargets } from './tui-targets.ts'

describe('TUI benchmark harness', () => {
	test('single scenario emits a benchmark case', async () => {
		const proc = Bun.spawn(
			[
				'bun',
				'run',
				'fixtures/benchmarks/tui.ts',
				'--scenario',
				'file-hub-first-paint',
				'--repeat',
				'1',
				'--warmup',
				'0',
				'--json',
			],
			{ stdout: 'pipe', stderr: 'pipe' },
		)
		const [stdout, stderr, code] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])
		expect(stderr).toBe('')
		expect(code).toBe(0)
		const suite = JSON.parse(stdout) as {
			cases: Array<{ name: string; metrics: { p95Ms: number } }>
		}
		expect(suite.cases[0]?.name).toBe('file-hub-first-paint')
		expect(suite.cases[0]?.metrics.p95Ms).toBeGreaterThanOrEqual(0)
	})

	test('target checker can skip scenarios not selected for a focused run', () => {
		const results = checkTuiTargets(
			createBenchmarkSuite({
				suite: 'tui',
				kind: 'synthetic',
				cases: [
					{
						name: 'paste-10k-cells',
						category: 'edit',
						dimensions: {},
						metrics: {
							repeat: 1,
							meanMs: 10,
							medianMs: 10,
							minMs: 10,
							maxMs: 10,
							p95Ms: 10,
							stddevMs: 0,
						},
						samples: [{ durationMs: 10 }],
					},
				],
			}),
			{ skipMissing: true },
		)
		expect(results.find((result) => result.target.scenario === 'paste-10k-cells')?.passed).toBe(
			true,
		)
		expect(results.some((result) => result.skipped)).toBe(true)
		expect(results.every((result) => result.passed)).toBe(true)
	})
})
