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
		'reports rows window first-row and window timings',
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
					'rows-window',
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
					readonly rowsWindowFirstRowMedianMs?: number
					readonly rowsWindowMedianMs?: number
					readonly rowsWindowCellsPerSecondMedian?: number
				}
			}
			expect(result.summary?.rowsWindowFirstRowMedianMs).toBeNumber()
			expect(result.summary?.rowsWindowMedianMs).toBeNumber()
			expect(result.summary?.rowsWindowCellsPerSecondMedian).toBeNumber()
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

	test(
		'reports full-fidelity read phase timings',
		async () => {
			const proc = Bun.spawn(
				[
					'bun',
					'run',
					'fixtures/benchmarks/xlsx-read-phase.ts',
					'--workload',
					'mixed-10pct-text',
					'--rows',
					'12',
					'--cols',
					'6',
					'--phase',
					'full-read',
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
					readonly fullReadXlsxMedianMs?: number
					readonly fullReadXlsxCellsPerSecondMedian?: number
				}
			}
			expect(result.summary?.fullReadXlsxMedianMs).toBeNumber()
			expect(result.summary?.fullReadXlsxCellsPerSecondMedian).toBeNumber()
		},
		{ timeout: 30_000 },
	)

	test(
		'infers input-file worksheet shape when rows and columns are omitted',
		async () => {
			const input = await buildWorkloadDataSet('mixed-10pct-text', 12, 6, 'raw-ooxml')
			const proc = Bun.spawn(
				[
					'bun',
					'run',
					'fixtures/benchmarks/xlsx-read-phase.ts',
					'--input-file',
					input.xlsxPath,
					'--phase',
					'full-read',
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
				readonly args?: {
					readonly rows?: number
					readonly cols?: number
				}
				readonly inputFileShape?: {
					readonly rows?: number
					readonly cols?: number
					readonly cells?: number
				}
				readonly summary?: {
					readonly fullReadXlsxMedianMs?: number
				}
			}
			expect(result.args?.rows).toBe(12)
			expect(result.args?.cols).toBe(6)
			expect(result.inputFileShape?.rows).toBe(12)
			expect(result.inputFileShape?.cols).toBe(6)
			expect(result.inputFileShape?.cells).toBe(72)
			expect(result.summary?.fullReadXlsxMedianMs).toBeNumber()
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
