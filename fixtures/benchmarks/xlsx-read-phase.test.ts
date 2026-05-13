import { describe, expect, test } from 'bun:test'
import { buildWorkloadDataSet } from './competitive-io.ts'
import { type Args, directOrderedReadAssertions } from './runners/ascend_runner.ts'

describe('xlsx-read-phase CLI', () => {
	test(
		'reports ZIP inflate and decode phase timings',
		async () => {
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
		},
		{ timeout: 30_000 },
	)

	test(
		'reports row streaming phase timings',
		async () => {
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
		},
		{ timeout: 30_000 },
	)

	test(
		'reports chunked row streaming phase timings',
		async () => {
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
		},
		{ timeout: 30_000 },
	)

	test(
		'reports capped first-window phase timings',
		async () => {
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
					'capped-agent-window',
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
					readonly cappedReadWindowMedianMs?: number
					readonly cappedAgentWindowMedianMs?: number
					readonly totalCappedAgentWindowMedianMs?: number
					readonly cappedReadWindowCellsPerSecondMedian?: number
				}
			}
			expect(result.summary?.cappedReadWindowMedianMs).toBeNumber()
			expect(result.summary?.cappedAgentWindowMedianMs).toBeNumber()
			expect(result.summary?.totalCappedAgentWindowMedianMs).toBeNumber()
			expect(result.summary?.cappedReadWindowCellsPerSecondMedian).toBeNumber()
		},
		{ timeout: 30_000 },
	)

	test('direct ordered verifier counts shared-string physical cells once', async () => {
		const input = await buildWorkloadDataSet('mixed-50pct-text', 8, 6, 'raw-ooxml')
		const args: Args = {
			operation: 'read',
			file: input.xlsxPath,
			mode: 'values',
			source: 'path',
			richMetadata: false,
			orderedHashes: true,
			streamOrderedHashes: false,
			repeat: 1,
			warmup: 0,
			json: true,
		}
		const assertions = await directOrderedReadAssertions(args, undefined)
		expect(assertions?.cellCount).toBe(48)
		expect(assertions?.physicalCellCount).toBe(48)
		expect(assertions?.firstPhysicalUsedRange).toBe(assertions?.firstUsedRange)
	})
})
