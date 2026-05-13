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
				'800',
				'--cols',
				'8',
				'--row-limit',
				'500',
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
				readonly tuiFirstPaintMedianMs?: number
				readonly tuiOpenMedianMs?: number
				readonly tuiRenderMedianMs?: number
				readonly tuiHydrateMedianMs?: number
				readonly cellsMedian?: number
				readonly payloadBytesMedian?: number
				readonly fullHydratedCellsMedian?: number
				readonly cappedHydratedCellsMedian?: number
				readonly tuiHydratedCellsMedian?: number
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
				readonly tuiOpenCallsMedian?: number
				readonly tuiHydratedOpenCountMedian?: number
				readonly tuiDocumentCacheHitsMedian?: number
				readonly apiPartial?: boolean
				readonly mcpPartial?: boolean
				readonly tuiPartial?: boolean
				readonly mcpPayloadBytesMedian?: number
				readonly tuiFrameBytesMedian?: number
			}
		}
		expect(payload.summary?.fullOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.cappedOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.apiFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.mcpFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.tuiFirstPaintMedianMs).toBeNumber()
		expect(payload.summary?.tuiOpenMedianMs).toBeNumber()
		expect(payload.summary?.tuiRenderMedianMs).toBeNumber()
		expect(payload.summary?.tuiHydrateMedianMs).toBeNumber()
		expect(payload.summary?.cellsMedian).toBe(500 * 8)
		expect(payload.summary?.payloadBytesMedian).toBeGreaterThan(0)
		expect(payload.summary?.fullHydratedCellsMedian).toBe(800 * 8)
		expect(payload.summary?.cappedHydratedCellsMedian).toBe(500 * 8)
		expect(payload.summary?.tuiHydratedCellsMedian).toBe(500 * 8)
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
		expect(payload.summary?.tuiOpenCallsMedian).toBe(1)
		expect(payload.summary?.tuiHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.tuiDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.apiPartial).toBe(true)
		expect(payload.summary?.mcpPartial).toBe(true)
		expect(payload.summary?.tuiPartial).toBe(true)
		expect(payload.summary?.mcpPayloadBytesMedian).toBeGreaterThan(0)
		expect(payload.summary?.tuiFrameBytesMedian).toBeGreaterThan(0)
	})
})
