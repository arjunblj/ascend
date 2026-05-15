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
				'--surface',
				'both',
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
				readonly preparedCommitVerifiedTotalStats?: {
					readonly sampleCount?: number
					readonly p95?: number
					readonly cv?: number
				}
				readonly preparedPlanStats?: {
					readonly sampleCount?: number
					readonly p95?: number
					readonly cv?: number
				}
				readonly preparedCommitPostWriteReopenStats?: {
					readonly sampleCount?: number
					readonly p95?: number
					readonly cv?: number
				}
				readonly preparedVerifyStats?: {
					readonly sampleCount?: number
					readonly p95?: number
					readonly cv?: number
				}
				readonly rssDeltaMbStats?: {
					readonly sampleCount?: number
					readonly p95?: number
					readonly cv?: number
				}
				readonly totalMedianMs?: number
				readonly fullTotalMedianMs?: number
				readonly preparedTotalMedianMs?: number
				readonly commitVerifiedTotalMedianMs?: number
				readonly fullCommitVerifiedTotalMedianMs?: number
				readonly preparedCommitVerifiedTotalMedianMs?: number
				readonly measuredSampleMedianMs?: number
				readonly readMedianMs?: number
				readonly commitPostWriteReopenMedianMs?: number
				readonly commitPostWriteCheckMedianMs?: number
				readonly commitPostWritePackageGraphAuditMedianMs?: number
				readonly preparedCommitPostWriteReopenMedianMs?: number
				readonly preparedCommitPostWriteCheckMedianMs?: number
				readonly preparedCommitPostWritePackageGraphAuditMedianMs?: number
				readonly payloadBytesMedian?: number
				readonly fullPayloadBytesMedian?: number
				readonly preparedPayloadBytesMedian?: number
				readonly commitVerifiedPayloadBytesMedian?: number
				readonly fullCommitVerifiedPayloadBytesMedian?: number
				readonly preparedCommitVerifiedPayloadBytesMedian?: number
				readonly planPayloadBytesMedian?: number
				readonly fullPlanPayloadBytesMedian?: number
				readonly preparedPlanPayloadBytesMedian?: number
				readonly compactWorkflowSpeedupVsFull?: number
				readonly commitVerifiedWorkflowSpeedupVsFull?: number
				readonly preparedWorkflowSpeedupVsCompact?: number
				readonly preparedCommitVerifiedWorkflowSpeedupVsCompact?: number
				readonly planPayloadReduction?: number
				readonly readCellsMedian?: number
				readonly readWindowRowsMedian?: number
				readonly planChangedCellCountMedian?: number
				readonly planEmittedChangedCellCountMedian?: number
				readonly preparedPlanChangedCellCountMedian?: number
				readonly preparedPlanEmittedChangedCellCountMedian?: number
				readonly compactHydratedOpenCountMedian?: number
				readonly commitVerifiedHydratedOpenCountMedian?: number
				readonly fullCommitVerifiedHydratedOpenCountMedian?: number
				readonly preparedCommitVerifiedHydratedOpenCountMedian?: number
				readonly fullHydratedOpenCountMedian?: number
				readonly preparedHydratedOpenCountMedian?: number
				readonly mcpTotalMedianMs?: number
				readonly mcpPreparedTotalMedianMs?: number
				readonly mcpCommitVerifiedTotalMedianMs?: number
				readonly mcpPreparedCommitVerifiedTotalMedianMs?: number
				readonly mcpPayloadBytesMedian?: number
				readonly mcpCommitVerifiedPayloadBytesMedian?: number
				readonly mcpReadCellsMedian?: number
				readonly mcpCompactHydratedOpenCountMedian?: number
				readonly mcpCommitVerifiedHydratedOpenCountMedian?: number
				readonly mcpPreparedHydratedOpenCountMedian?: number
				readonly mcpPreparedCommitVerifiedHydratedOpenCountMedian?: number
				readonly mcpReadPartial?: boolean
				readonly mcpValid?: boolean
				readonly mcpPreparedValid?: boolean
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
		expect(payload.summary?.commitVerifiedTotalMedianMs).toBeNumber()
		expect(payload.summary?.fullCommitVerifiedTotalMedianMs).toBeNumber()
		expect(payload.summary?.preparedCommitVerifiedTotalMedianMs).toBeNumber()
		expect(payload.summary?.preparedCommitVerifiedTotalStats?.sampleCount).toBe(1)
		expect(payload.summary?.preparedCommitVerifiedTotalStats?.p95).toBeNumber()
		expect(payload.summary?.preparedCommitVerifiedTotalStats?.cv).toBeNumber()
		expect(payload.summary?.measuredSampleMedianMs).toBeNumber()
		expect(payload.summary?.readMedianMs).toBeNumber()
		expect(payload.summary?.commitPostWriteReopenMedianMs).toBeNumber()
		expect(payload.summary?.commitPostWriteCheckMedianMs).toBeNumber()
		expect(payload.summary?.commitPostWritePackageGraphAuditMedianMs).toBeNumber()
		expect(payload.summary?.preparedCommitPostWriteReopenMedianMs).toBeNumber()
		expect(payload.summary?.preparedCommitPostWriteReopenStats?.sampleCount).toBe(1)
		expect(payload.summary?.preparedCommitPostWriteReopenStats?.p95).toBeNumber()
		expect(payload.summary?.preparedCommitPostWriteCheckMedianMs).toBeNumber()
		expect(payload.summary?.preparedCommitPostWritePackageGraphAuditMedianMs).toBeNumber()
		expect(payload.summary?.payloadBytesMedian).toBeNumber()
		expect(payload.summary?.fullPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.preparedPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.commitVerifiedPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.fullCommitVerifiedPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.preparedCommitVerifiedPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.planPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.fullPlanPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.preparedPlanPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.preparedPlanStats?.sampleCount).toBe(1)
		expect(payload.summary?.preparedPlanStats?.p95).toBeNumber()
		expect(payload.summary?.preparedVerifyStats?.sampleCount).toBe(1)
		expect(payload.summary?.preparedVerifyStats?.p95).toBeNumber()
		expect(payload.summary?.compactWorkflowSpeedupVsFull).toBeNumber()
		expect(payload.summary?.commitVerifiedWorkflowSpeedupVsFull).toBeNumber()
		expect(payload.summary?.preparedWorkflowSpeedupVsCompact).toBeNumber()
		expect(payload.summary?.preparedCommitVerifiedWorkflowSpeedupVsCompact).toBeNumber()
		expect(payload.summary?.planPayloadReduction).toBeNumber()
		expect(payload.summary?.readCellsMedian).toBe(60)
		expect(payload.summary?.readWindowRowsMedian).toBe(10)
		expect(payload.summary?.planChangedCellCountMedian).toBe(12)
		expect(payload.summary?.planEmittedChangedCellCountMedian).toBe(12)
		expect(payload.summary?.preparedPlanChangedCellCountMedian).toBe(12)
		expect(payload.summary?.preparedPlanEmittedChangedCellCountMedian).toBe(12)
		expect(payload.summary?.compactHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.commitVerifiedHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.fullCommitVerifiedHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.preparedCommitVerifiedHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.fullHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.preparedHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.mcpTotalMedianMs).toBeNumber()
		expect(payload.summary?.mcpPreparedTotalMedianMs).toBeNumber()
		expect(payload.summary?.mcpCommitVerifiedTotalMedianMs).toBeNumber()
		expect(payload.summary?.mcpPreparedCommitVerifiedTotalMedianMs).toBeNumber()
		expect(payload.summary?.mcpPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.mcpCommitVerifiedPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.mcpReadCellsMedian).toBe(60)
		expect(payload.summary?.mcpCompactHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.mcpCommitVerifiedHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.mcpPreparedHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.mcpPreparedCommitVerifiedHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.planHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.preparedPlanHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.commitHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.preparedCommitHydratedOpenCountMedian).toBeNumber()
		expect(payload.summary?.documentCacheHitCountMedian).toBeNumber()
		expect(payload.summary?.preparedHydratedOpenCountMedian).toBeLessThan(
			payload.summary?.compactHydratedOpenCountMedian ?? 0,
		)
		expect(payload.summary?.commitVerifiedHydratedOpenCountMedian).toBeLessThan(
			payload.summary?.compactHydratedOpenCountMedian ?? 0,
		)
		expect(payload.summary?.preparedCommitVerifiedHydratedOpenCountMedian).toBeLessThanOrEqual(
			payload.summary?.preparedHydratedOpenCountMedian ?? 0,
		)
		expect(payload.summary?.mcpCommitVerifiedHydratedOpenCountMedian).toBeLessThan(
			payload.summary?.mcpCompactHydratedOpenCountMedian ?? 0,
		)
		expect(payload.summary?.mcpPreparedCommitVerifiedHydratedOpenCountMedian).toBeLessThan(
			payload.summary?.mcpPreparedHydratedOpenCountMedian ?? 0,
		)
		expect(payload.summary?.preparedCommitHydratedOpenCountMedian).toBeLessThan(
			payload.summary?.commitHydratedOpenCountMedian ?? 0,
		)
		expect(payload.summary?.mutationCountMedian).toBe(12)
		expect(payload.summary?.rssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.rssDeltaMbStats?.sampleCount).toBe(1)
		expect(payload.summary?.rssDeltaMbStats?.p95).toBeNumber()
		expect(payload.summary?.readPartial).toBe(true)
		expect(payload.summary?.valid).toBe(true)
		expect(payload.summary?.preparedValid).toBe(true)
		expect(payload.summary?.mcpReadPartial).toBe(true)
		expect(payload.summary?.mcpValid).toBe(true)
		expect(payload.summary?.mcpPreparedValid).toBe(true)
	})

	test('runs the full workflow against an existing input workbook without deleting it', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--input-file',
				'fixtures/xlsx/poi/WithVariousData.xlsx',
				'--row-limit',
				'5',
				'--mutations',
				'1',
				'--approval',
				'all',
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
			xlsxPath: 'fixtures/xlsx/poi/WithVariousData.xlsx',
			range: 'A1:C65536',
			sheet: 'Sheet1',
			rows: 65_536,
			cols: 3,
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
