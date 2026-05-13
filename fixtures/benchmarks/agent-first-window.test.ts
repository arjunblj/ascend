import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./agent-first-window.ts', import.meta.url))

describe('agent first-window benchmark', () => {
	test('reports full, capped, API, and MCP first-window timings with partial load metadata', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--rows',
				'120',
				'--cols',
				'8',
				'--row-limit',
				'25',
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
				readonly fullOpenWindowMedianMs?: number
				readonly cappedOpenWindowMedianMs?: number
				readonly apiFirstWindowMedianMs?: number
				readonly mcpFirstWindowMedianMs?: number
				readonly fullHydratedCellsMedian?: number
				readonly cappedHydratedCellsMedian?: number
				readonly fullOpenCallsMedian?: number
				readonly fullHydratedOpenCountMedian?: number
				readonly fullDocumentCacheHitsMedian?: number
				readonly cappedOpenCallsMedian?: number
				readonly cappedHydratedOpenCountMedian?: number
				readonly cappedDocumentCacheHitsMedian?: number
				readonly apiOpenCallsMedian?: number
				readonly apiHydratedOpenCountMedian?: number
				readonly apiDocumentCacheHitsMedian?: number
				readonly mcpOpenCallsMedian?: number
				readonly mcpHydratedOpenCountMedian?: number
				readonly mcpDocumentCacheHitsMedian?: number
				readonly apiPartial?: boolean
				readonly mcpPartial?: boolean
				readonly mcpPayloadBytesMedian?: number
			}
		}
		expect(payload.summary?.fullOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.cappedOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.apiFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.mcpFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.fullHydratedCellsMedian).toBe(120 * 8)
		expect(payload.summary?.cappedHydratedCellsMedian).toBe(25 * 8)
		expect(payload.summary?.fullOpenCallsMedian).toBe(1)
		expect(payload.summary?.fullHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.fullDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.cappedOpenCallsMedian).toBe(1)
		expect(payload.summary?.cappedHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.cappedDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.apiOpenCallsMedian).toBe(1)
		expect(payload.summary?.apiHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.apiDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.mcpOpenCallsMedian).toBe(1)
		expect(payload.summary?.mcpHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.mcpDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.apiPartial).toBe(true)
		expect(payload.summary?.mcpPartial).toBe(true)
		expect(payload.summary?.mcpPayloadBytesMedian).toBeGreaterThan(0)
	})
})
