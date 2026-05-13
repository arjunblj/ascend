import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./tui-first-paint.ts', import.meta.url))

describe('TUI first-paint benchmark', () => {
	test('compares full open paint against a row-limited preview paint', async () => {
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
				readonly fullTotalMedianMs?: number
				readonly previewTotalMedianMs?: number
				readonly fullFrameBytesMedian?: number
				readonly previewFrameBytesMedian?: number
				readonly fullHydratedCellsMedian?: number
				readonly previewHydratedCellsMedian?: number
				readonly previewPartial?: boolean
				readonly previewReadOnly?: boolean
				readonly fullPartial?: boolean
			}
		}
		expect(payload.summary?.fullTotalMedianMs).toBeNumber()
		expect(payload.summary?.previewTotalMedianMs).toBeNumber()
		expect(payload.summary?.fullFrameBytesMedian).toBeGreaterThan(0)
		expect(payload.summary?.previewFrameBytesMedian).toBeGreaterThan(0)
		expect(payload.summary?.fullHydratedCellsMedian).toBe(480)
		expect(payload.summary?.previewHydratedCellsMedian).toBe(60)
		expect(payload.summary?.previewPartial).toBe(true)
		expect(payload.summary?.previewReadOnly).toBe(true)
		expect(payload.summary?.fullPartial).toBe(false)
	})
})
