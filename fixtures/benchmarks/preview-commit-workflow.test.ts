import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./preview-commit-workflow.ts', import.meta.url))

describe('preview commit workflow benchmark', () => {
	test('compares direct full, preview, and preview-promotion commit paths', async () => {
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
				readonly directFullTotalMedianMs?: number
				readonly previewTotalMedianMs?: number
				readonly previewPromoteTotalMedianMs?: number
				readonly directFullHydratedCellsMedian?: number
				readonly previewHydratedCellsMedian?: number
				readonly promotedHydratedCellsMedian?: number
				readonly windowCellsMedian?: number
				readonly valid?: boolean
			}
		}
		expect(payload.summary?.directFullTotalMedianMs).toBeNumber()
		expect(payload.summary?.previewTotalMedianMs).toBeNumber()
		expect(payload.summary?.previewPromoteTotalMedianMs).toBeNumber()
		expect(payload.summary?.directFullHydratedCellsMedian).toBe(480)
		expect(payload.summary?.previewHydratedCellsMedian).toBe(60)
		expect(payload.summary?.promotedHydratedCellsMedian).toBe(480)
		expect(payload.summary?.windowCellsMedian).toBe(60)
		expect(payload.summary?.valid).toBe(true)
	})
})
