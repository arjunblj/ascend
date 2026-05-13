import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./agent-workflow.ts', import.meta.url))

describe('agent workflow benchmark', () => {
	test('measures inspect, first-window read, plan, commit, and verify loop', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'80',
				'--cols',
				'6',
				'--row-limit',
				'10',
				'--mutations',
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
				readonly totalMedianMs?: number
				readonly fullTotalMedianMs?: number
				readonly preparedTotalMedianMs?: number
				readonly measuredSampleMedianMs?: number
				readonly readMedianMs?: number
				readonly payloadBytesMedian?: number
				readonly fullPayloadBytesMedian?: number
				readonly preparedPayloadBytesMedian?: number
				readonly planPayloadBytesMedian?: number
				readonly fullPlanPayloadBytesMedian?: number
				readonly preparedPlanPayloadBytesMedian?: number
				readonly compactWorkflowSpeedupVsFull?: number
				readonly preparedWorkflowSpeedupVsCompact?: number
				readonly planPayloadReduction?: number
				readonly readCellsMedian?: number
				readonly readWindowRowsMedian?: number
				readonly planChangedCellCountMedian?: number
				readonly planEmittedChangedCellCountMedian?: number
				readonly preparedPlanChangedCellCountMedian?: number
				readonly preparedPlanEmittedChangedCellCountMedian?: number
				readonly compactHydratedOpenCountMedian?: number
				readonly fullHydratedOpenCountMedian?: number
				readonly preparedHydratedOpenCountMedian?: number
				readonly planHydratedOpenCountMedian?: number
				readonly preparedPlanHydratedOpenCountMedian?: number
				readonly commitHydratedOpenCountMedian?: number
				readonly preparedCommitHydratedOpenCountMedian?: number
				readonly documentCacheHitCountMedian?: number
				readonly mutationCountMedian?: number
				readonly rssDeltaMbMedian?: number
				readonly readPartial?: boolean
				readonly valid?: boolean
				readonly preparedValid?: boolean
			}
		}
		expect(payload.summary?.totalMedianMs).toBeNumber()
		expect(payload.summary?.fullTotalMedianMs).toBeNumber()
		expect(payload.summary?.preparedTotalMedianMs).toBeNumber()
		expect(payload.summary?.measuredSampleMedianMs).toBeNumber()
		expect(payload.summary?.readMedianMs).toBeNumber()
		expect(payload.summary?.payloadBytesMedian).toBeNumber()
		expect(payload.summary?.fullPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.preparedPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.planPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.fullPlanPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.preparedPlanPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.compactWorkflowSpeedupVsFull).toBeNumber()
		expect(payload.summary?.preparedWorkflowSpeedupVsCompact).toBeNumber()
		expect(payload.summary?.planPayloadReduction).toBeNumber()
		expect(payload.summary?.readCellsMedian).toBe(60)
		expect(payload.summary?.readWindowRowsMedian).toBe(10)
		expect(payload.summary?.planChangedCellCountMedian).toBe(12)
		expect(payload.summary?.planEmittedChangedCellCountMedian).toBe(12)
		expect(payload.summary?.preparedPlanChangedCellCountMedian).toBe(12)
		expect(payload.summary?.preparedPlanEmittedChangedCellCountMedian).toBe(12)
		expect(payload.summary?.compactHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.fullHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.preparedHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.planHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.preparedPlanHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.commitHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.preparedCommitHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.documentCacheHitCountMedian).toBeNumber()
		expect(payload.summary?.preparedHydratedOpenCountMedian).toBeLessThan(
			payload.summary?.compactHydratedOpenCountMedian ?? 0,
		)
		expect(payload.summary?.preparedCommitHydratedOpenCountMedian).toBeLessThan(
			payload.summary?.commitHydratedOpenCountMedian ?? 0,
		)
		expect(payload.summary?.mutationCountMedian).toBe(12)
		expect(payload.summary?.rssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.readPartial).toBe(true)
		expect(payload.summary?.valid).toBe(true)
		expect(payload.summary?.preparedValid).toBe(true)
	})

	test('runs the full workflow against an existing input workbook without deleting it', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--input-file',
				'fixtures/xlsx/poi/SampleSS.xlsx',
				'--row-limit',
				'5',
				'--mutations',
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
				readonly range?: string
				readonly sheet?: string
				readonly rows?: number
				readonly cols?: number
				readonly cleanup?: boolean
				readonly source?: string
			}
			readonly summary?: {
				readonly readCellsMedian?: number
				readonly readWindowRowsMedian?: number
				readonly mutationCountMedian?: number
				readonly readPartial?: boolean
				readonly valid?: boolean
				readonly preparedValid?: boolean
			}
		}
		expect(payload.input).toEqual({
			xlsxPath: 'fixtures/xlsx/poi/SampleSS.xlsx',
			range: 'A1:J65536',
			sheet: 'First Sheet',
			rows: 65_536,
			cols: 10,
			cleanup: false,
			source: 'input-file',
		})
		expect(payload.summary?.readCellsMedian).toBeGreaterThan(0)
		expect(payload.summary?.readWindowRowsMedian).toBe(5)
		expect(payload.summary?.mutationCountMedian).toBe(1)
		expect(payload.summary?.readPartial).toBe(true)
		expect(payload.summary?.valid).toBe(true)
		expect(payload.summary?.preparedValid).toBe(true)
	}, 20_000)
})
