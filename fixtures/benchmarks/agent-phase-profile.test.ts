import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./agent-phase-profile.ts', import.meta.url))

describe('agent phase profile benchmark', () => {
	test('reports plan and commit phase timing from workflow progress events', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'80',
				'--cols',
				'6',
				'--updates',
				'12',
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
		expect(exitCode, stderr).toBe(0)
		const payload = JSON.parse(stdout) as {
			readonly summary?: {
				readonly planMedianMs?: number
				readonly commitMedianMs?: number
				readonly sharedPlanMedianMs?: number
				readonly sharedCommitMedianMs?: number
				readonly sharedTotalMedianMs?: number
				readonly sharedWorkflowSpeedupVsCold?: number
				readonly commitOutputBytesMedian?: number
				readonly sharedCommitOutputBytesMedian?: number
				readonly operationCountMedian?: number
				readonly updateCountMedian?: number
				readonly changedCellsMedian?: number
				readonly commitChangedCellsMedian?: number
				readonly sharedChangedCellsMedian?: number
				readonly sharedCommitChangedCellsMedian?: number
				readonly postWriteValid?: boolean
				readonly sharedPostWriteValid?: boolean
				readonly planPhaseMedianMs?: Record<string, number>
				readonly commitPhaseMedianMs?: Record<string, number>
				readonly sharedPlanPhaseMedianMs?: Record<string, number>
				readonly sharedCommitPhaseMedianMs?: Record<string, number>
				readonly sharedPlanPhaseStats?: Record<
					string,
					{ readonly sampleCount?: number; readonly p95?: number; readonly cv?: number }
				>
				readonly sharedCommitPhaseStats?: Record<
					string,
					{ readonly sampleCount?: number; readonly p95?: number; readonly cv?: number }
				>
				readonly commitTimingMedianMs?: Record<string, number>
				readonly sharedCommitTimingMedianMs?: Record<string, number>
			}
		}
		expect(payload.summary?.planMedianMs).toBeNumber()
		expect(payload.summary?.commitMedianMs).toBeNumber()
		expect(payload.summary?.sharedPlanMedianMs).toBeNumber()
		expect(payload.summary?.sharedCommitMedianMs).toBeNumber()
		expect(payload.summary?.sharedTotalMedianMs).toBeNumber()
		expect(payload.summary?.sharedWorkflowSpeedupVsCold).toBeNumber()
		expect(payload.summary?.commitOutputBytesMedian).toBeGreaterThan(100)
		expect(payload.summary?.sharedCommitOutputBytesMedian).toBeGreaterThan(100)
		expect(payload.summary?.operationCountMedian).toBe(1)
		expect(payload.summary?.updateCountMedian).toBe(12)
		expect(payload.summary?.changedCellsMedian).toBe(12)
		expect(payload.summary?.commitChangedCellsMedian).toBe(12)
		expect(payload.summary?.sharedChangedCellsMedian).toBe(12)
		expect(payload.summary?.sharedCommitChangedCellsMedian).toBe(12)
		expect(payload.summary?.postWriteValid).toBe(true)
		expect(payload.summary?.sharedPostWriteValid).toBe(true)
		expect(payload.summary?.planPhaseMedianMs?.['load-workbook']).toBeNumber()
		expect(payload.summary?.planPhaseMedianMs?.preview).toBeNumber()
		expect(payload.summary?.commitPhaseMedianMs?.write).toBeNumber()
		expect(payload.summary?.commitPhaseMedianMs?.['post-write']).toBeNumber()
		expect(payload.summary?.sharedPlanPhaseMedianMs?.preview).toBeNumber()
		expect(payload.summary?.sharedCommitPhaseMedianMs?.write).toBeNumber()
		expect(payload.summary?.sharedPlanPhaseStats?.preview).toMatchObject({
			sampleCount: 1,
			p95: payload.summary?.sharedPlanPhaseMedianMs?.preview,
			cv: 0,
		})
		expect(payload.summary?.sharedCommitPhaseStats?.write).toMatchObject({
			sampleCount: 1,
			p95: payload.summary?.sharedCommitPhaseMedianMs?.write,
			cv: 0,
		})
		expect(payload.summary?.commitTimingMedianMs?.writePolicyCheckMs).toBeNumber()
		expect(payload.summary?.commitTimingMedianMs?.toBytesMs).toBeNumber()
		expect(payload.summary?.sharedCommitTimingMedianMs?.writePolicyBuildMs).toBeNumber()
		expect(payload.summary?.sharedCommitTimingMedianMs?.outputHashMs).toBeNumber()
	})

	test('profiles agent phases against an existing input workbook without deleting it', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--input-file',
				'fixtures/xlsx/poi/SampleSS.xlsx',
				'--updates',
				'1',
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
		expect(exitCode, stderr).toBe(0)
		const payload = JSON.parse(stdout) as {
			readonly input?: {
				readonly xlsxPath?: string
				readonly sheet?: string
				readonly rows?: number
				readonly cols?: number
				readonly cleanup?: boolean
				readonly source?: string
			}
			readonly summary?: {
				readonly sharedTotalMedianMs?: number
				readonly updateCountMedian?: number
				readonly commitOutputBytesMedian?: number
				readonly sharedCommitOutputBytesMedian?: number
				readonly sharedChangedCellsMedian?: number
				readonly postWriteValid?: boolean
				readonly sharedPostWriteValid?: boolean
				readonly sharedCommitPhaseMedianMs?: Record<string, number>
				readonly sharedCommitTimingMedianMs?: Record<string, number>
			}
		}
		expect(payload.input).toEqual({
			xlsxPath: 'fixtures/xlsx/poi/SampleSS.xlsx',
			sheet: 'First Sheet',
			rows: 65_536,
			cols: 2,
			cleanup: false,
			source: 'input-file',
		})
		expect(payload.summary?.sharedTotalMedianMs).toBeNumber()
		expect(payload.summary?.updateCountMedian).toBe(1)
		expect(payload.summary?.commitOutputBytesMedian).toBeGreaterThan(100)
		expect(payload.summary?.sharedCommitOutputBytesMedian).toBeGreaterThan(100)
		expect(payload.summary?.sharedChangedCellsMedian).toBe(1)
		expect(payload.summary?.postWriteValid).toBe(true)
		expect(payload.summary?.sharedPostWriteValid).toBe(true)
		expect(payload.summary?.sharedCommitPhaseMedianMs?.write).toBeNumber()
		expect(payload.summary?.sharedCommitTimingMedianMs?.toBytesMs).toBeNumber()
	}, 20_000)

	test('can stream progress phase events while enforcing a timeout guard', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'40',
				'--cols',
				'4',
				'--updates',
				'2',
				'--repeat',
				'1',
				'--warmup',
				'0',
				'--timeout-ms',
				'300000',
				'--progress',
				'--json',
			],
			{ cwd: process.cwd(), stderr: 'pipe', stdout: 'pipe' },
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])
		expect(exitCode, stderr).toBe(0)
		const payload = JSON.parse(stdout) as {
			readonly summary?: {
				readonly postWriteValid?: boolean
				readonly sharedPostWriteValid?: boolean
			}
		}
		const events = stderr
			.trim()
			.split('\n')
			.filter(Boolean)
			.map((line) => JSON.parse(line) as { readonly sample?: number; readonly phase?: string })
		expect(events).toContainEqual(expect.objectContaining({ sample: 1 }))
		expect(events).toContainEqual(expect.objectContaining({ sample: 1, phase: 'write' }))
		expect(events).toContainEqual(expect.objectContaining({ sample: 1, phase: 'post-write' }))
		expect(payload.summary?.postWriteValid).toBe(true)
		expect(payload.summary?.sharedPostWriteValid).toBe(true)
	}, 20_000)
})
