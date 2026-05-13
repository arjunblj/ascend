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
				readonly readMedianMs?: number
				readonly payloadBytesMedian?: number
				readonly planPayloadBytesMedian?: number
				readonly readCellsMedian?: number
				readonly readWindowRowsMedian?: number
				readonly mutationCountMedian?: number
				readonly rssDeltaMbMedian?: number
				readonly readPartial?: boolean
				readonly valid?: boolean
			}
		}
		expect(payload.summary?.totalMedianMs).toBeNumber()
		expect(payload.summary?.readMedianMs).toBeNumber()
		expect(payload.summary?.payloadBytesMedian).toBeNumber()
		expect(payload.summary?.planPayloadBytesMedian).toBeNumber()
		expect(payload.summary?.readCellsMedian).toBe(60)
		expect(payload.summary?.readWindowRowsMedian).toBe(10)
		expect(payload.summary?.mutationCountMedian).toBe(1)
		expect(payload.summary?.rssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.readPartial).toBe(true)
		expect(payload.summary?.valid).toBe(true)
	})
})
