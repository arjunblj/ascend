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
				readonly operationCountMedian?: number
				readonly updateCountMedian?: number
				readonly changedCellsMedian?: number
				readonly commitChangedCellsMedian?: number
				readonly postWriteValid?: boolean
				readonly planPhaseMedianMs?: Record<string, number>
				readonly commitPhaseMedianMs?: Record<string, number>
			}
		}
		expect(payload.summary?.planMedianMs).toBeNumber()
		expect(payload.summary?.commitMedianMs).toBeNumber()
		expect(payload.summary?.operationCountMedian).toBe(1)
		expect(payload.summary?.updateCountMedian).toBe(12)
		expect(payload.summary?.changedCellsMedian).toBe(12)
		expect(payload.summary?.commitChangedCellsMedian).toBe(12)
		expect(payload.summary?.postWriteValid).toBe(true)
		expect(payload.summary?.planPhaseMedianMs?.['load-workbook']).toBeNumber()
		expect(payload.summary?.planPhaseMedianMs?.preview).toBeNumber()
		expect(payload.summary?.commitPhaseMedianMs?.write).toBeNumber()
		expect(payload.summary?.commitPhaseMedianMs?.['post-write']).toBeNumber()
	})
})
