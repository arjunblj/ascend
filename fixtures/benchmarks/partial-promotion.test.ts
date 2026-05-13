import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./partial-promotion.ts', import.meta.url))

describe('partial promotion benchmark', () => {
	test('measures first-window preview and later full promotion', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'80',
				'--cols',
				'6',
				'--preview-rows',
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
				readonly directFullTotalMedianMs?: number
				readonly previewFirstWindowMedianMs?: number
				readonly promoteMedianMs?: number
				readonly directFullHydratedCellsMedian?: number
				readonly previewHydratedCellsMedian?: number
				readonly promotedHydratedCellsMedian?: number
				readonly firstWindowCellsMedian?: number
				readonly previewPartial?: boolean
				readonly promotedPartial?: boolean
			}
		}
		expect(payload.summary?.directFullTotalMedianMs).toBeNumber()
		expect(payload.summary?.previewFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.promoteMedianMs).toBeNumber()
		expect(payload.summary?.directFullHydratedCellsMedian).toBe(480)
		expect(payload.summary?.previewHydratedCellsMedian).toBe(60)
		expect(payload.summary?.promotedHydratedCellsMedian).toBe(480)
		expect(payload.summary?.firstWindowCellsMedian).toBe(60)
		expect(payload.summary?.previewPartial).toBe(true)
		expect(payload.summary?.promotedPartial).toBe(false)
	})
})
