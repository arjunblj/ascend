import { describe, expect, test } from 'bun:test'

describe('xlsx-write-phase CLI', () => {
	test('runs workbook streaming path for sparse workloads', async () => {
		const proc = Bun.spawn(
			[
				'bun',
				'run',
				'fixtures/benchmarks/xlsx-write-phase.ts',
				'--workload',
				'sparse-wide',
				'--rows',
				'10',
				'--cols',
				'5',
				'--streaming',
				'--repeat',
				'1',
				'--warmup',
				'0',
				'--json',
			],
			{ cwd: process.cwd(), stderr: 'pipe', stdout: 'pipe' },
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])
		expect(stderr).toBe('')
		expect(exitCode).toBe(0)
		const result = JSON.parse(stdout) as { summary?: { writerPath?: string } }
		expect(result.summary?.writerPath).toBe('workbook-streaming')
	})
})
