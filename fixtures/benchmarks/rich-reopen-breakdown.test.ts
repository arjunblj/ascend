import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./rich-reopen-breakdown.ts', import.meta.url))

describe('rich reopen breakdown benchmark', () => {
	test('measures workbook open modes used around post-write verification', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'80',
				'--cols',
				'6',
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
				readonly modes?: Record<
					string,
					{
						readonly totalMedianMs?: number
						readonly cellCountMedian?: number
						readonly isPartial?: boolean
						readonly cellsHydrated?: boolean
						readonly richSheetMetadataHydrated?: boolean
					}
				>
				readonly ratios?: {
					readonly fullRichOverFull?: number
					readonly fullOverValuesCapped500?: number
				}
			}
		}
		const modes = payload.summary?.modes
		expect(modes?.metadataOnly?.totalMedianMs).toBeNumber()
		expect(modes?.valuesCapped500?.totalMedianMs).toBeNumber()
		expect(modes?.valuesFull?.totalMedianMs).toBeNumber()
		expect(modes?.valuesFullRich?.totalMedianMs).toBeNumber()
		expect(modes?.formulaFull?.totalMedianMs).toBeNumber()
		expect(modes?.full?.totalMedianMs).toBeNumber()
		expect(modes?.fullRich?.totalMedianMs).toBeNumber()
		expect(modes?.metadataOnly?.cellCountMedian).toBe(0)
		expect(modes?.metadataOnly?.isPartial).toBe(true)
		expect(modes?.valuesCapped500?.isPartial).toBe(true)
		expect(modes?.valuesCapped500?.cellsHydrated).toBe(true)
		expect(modes?.valuesFull?.isPartial).toBe(true)
		expect(modes?.valuesFull?.cellCountMedian).toBe(480)
		expect(modes?.valuesFullRich?.richSheetMetadataHydrated).toBe(true)
		expect(modes?.full?.isPartial).toBe(false)
		expect(modes?.full?.cellCountMedian).toBe(480)
		expect(modes?.fullRich?.isPartial).toBe(false)
		expect(payload.summary?.ratios?.fullRichOverFull).toBeNumber()
		expect(payload.summary?.ratios?.fullOverValuesCapped500).toBeNumber()
	})
})
