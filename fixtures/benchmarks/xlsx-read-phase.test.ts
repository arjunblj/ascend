import { describe, expect, test } from 'bun:test'

describe('xlsx-read-phase CLI', () => {
	test('reports ZIP inflate and decode phase timings', async () => {
		const proc = Bun.spawn(
			[
				'bun',
				'run',
				'fixtures/benchmarks/xlsx-read-phase.ts',
				'--workload',
				'mixed-50pct-text',
				'--rows',
				'12',
				'--cols',
				'6',
				'--phase',
				'zip',
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
		const result = JSON.parse(stdout) as {
			readonly summary?: {
				readonly zipOpenMedianMs?: number
				readonly worksheetInflateMedianMs?: number
				readonly worksheetDecodeMedianMs?: number
				readonly dominantPhase?: string
			}
		}
		expect(result.summary?.zipOpenMedianMs).toBeNumber()
		expect(result.summary?.worksheetInflateMedianMs).toBeNumber()
		expect(result.summary?.worksheetDecodeMedianMs).toBeNumber()
		expect(result.summary?.dominantPhase).toBeString()
	})

	test('reports row streaming phase timings', async () => {
		const proc = Bun.spawn(
			[
				'bun',
				'run',
				'fixtures/benchmarks/xlsx-read-phase.ts',
				'--workload',
				'mixed-50pct-text',
				'--rows',
				'12',
				'--cols',
				'6',
				'--phase',
				'rows',
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
		const result = JSON.parse(stdout) as {
			readonly summary?: {
				readonly rowsStreamMedianMs?: number
				readonly rowsStreamCellsPerSecondMedian?: number
			}
		}
		expect(result.summary?.rowsStreamMedianMs).toBeNumber()
		expect(result.summary?.rowsStreamCellsPerSecondMedian).toBeNumber()
	})

	test('reports chunked row streaming phase timings', async () => {
		const proc = Bun.spawn(
			[
				'bun',
				'run',
				'fixtures/benchmarks/xlsx-read-phase.ts',
				'--workload',
				'mixed-50pct-text',
				'--rows',
				'12',
				'--cols',
				'6',
				'--phase',
				'rows-chunked',
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
		const result = JSON.parse(stdout) as {
			readonly summary?: {
				readonly rowsStreamChunkedMedianMs?: number
				readonly rowsStreamChunkedCellsPerSecondMedian?: number
			}
		}
		expect(result.summary?.rowsStreamChunkedMedianMs).toBeNumber()
		expect(result.summary?.rowsStreamChunkedCellsPerSecondMedian).toBeNumber()
	})
})
