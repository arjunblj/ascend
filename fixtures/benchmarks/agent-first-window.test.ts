import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'

const runnerPath = fileURLToPath(new URL('./agent-first-window.ts', import.meta.url))

describe('agent first-window benchmark', () => {
	test('reports full, capped, CLI, API, and MCP first-window timings with partial load metadata', async () => {
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
				readonly fullWarmOpenWindowMedianMs?: number
				readonly cappedOpenWindowMedianMs?: number
				readonly cappedWarmOpenWindowMedianMs?: number
				readonly apiFirstWindowMedianMs?: number
				readonly apiWarmFirstWindowMedianMs?: number
				readonly apiCompactDefaultMedianMs?: number
				readonly apiWarmCompactDefaultMedianMs?: number
				readonly cliReadFirstWindowMedianMs?: number
				readonly cliWarmReadFirstWindowMedianMs?: number
				readonly mcpFirstWindowMedianMs?: number
				readonly mcpWarmFirstWindowMedianMs?: number
				readonly mcpCompactDefaultMedianMs?: number
				readonly mcpWarmCompactDefaultMedianMs?: number
				readonly tuiFirstPaintMedianMs?: number
				readonly tuiWarmFirstPaintMedianMs?: number
				readonly tuiOpenMedianMs?: number
				readonly tuiWarmOpenMedianMs?: number
				readonly tuiRenderMedianMs?: number
				readonly tuiWarmRenderMedianMs?: number
				readonly tuiHydrateMedianMs?: number
				readonly tuiWarmHydrateMedianMs?: number
				readonly fullRssDeltaMbMedian?: number
				readonly fullRetainedRssDeltaMbMedian?: number
				readonly cappedRssDeltaMbMedian?: number
				readonly cappedRetainedRssDeltaMbMedian?: number
				readonly apiRssDeltaMbMedian?: number
				readonly apiRetainedRssDeltaMbMedian?: number
				readonly mcpRssDeltaMbMedian?: number
				readonly mcpRetainedRssDeltaMbMedian?: number
				readonly tuiRssDeltaMbMedian?: number
				readonly tuiRetainedRssDeltaMbMedian?: number
				readonly cellsMedian?: number
				readonly payloadBytesMedian?: number
				readonly apiCompactDefaultPayloadBytesMedian?: number
				readonly apiCompactDefaultCellsMedian?: number
				readonly fullHydratedCellsMedian?: number
				readonly cappedHydratedCellsMedian?: number
				readonly tuiHydratedCellsMedian?: number
				readonly fullOpenCallsMedian?: number
				readonly fullWarmOpenCallsMedian?: number
				readonly fullHydratedOpenCountMedian?: number
				readonly fullWarmHydratedOpenCountMedian?: number
				readonly fullDocumentCacheHitsMedian?: number
				readonly fullWarmDocumentCacheHitsMedian?: number
				readonly cappedOpenCallsMedian?: number
				readonly cappedWarmOpenCallsMedian?: number
				readonly cappedHydratedOpenCountMedian?: number
				readonly cappedWarmHydratedOpenCountMedian?: number
				readonly cappedDocumentCacheHitsMedian?: number
				readonly cappedWarmDocumentCacheHitsMedian?: number
				readonly apiOpenCallsMedian?: number
				readonly apiWarmOpenCallsMedian?: number
				readonly apiHydratedOpenCountMedian?: number
				readonly apiWarmHydratedOpenCountMedian?: number
				readonly apiDocumentCacheHitsMedian?: number
				readonly apiWarmDocumentCacheHitsMedian?: number
				readonly cliOpenCallsMedian?: number
				readonly cliWarmOpenCallsMedian?: number
				readonly cliHydratedOpenCountMedian?: number
				readonly cliWarmHydratedOpenCountMedian?: number
				readonly cliDocumentCacheHitsMedian?: number
				readonly cliWarmDocumentCacheHitsMedian?: number
				readonly mcpOpenCallsMedian?: number
				readonly mcpWarmOpenCallsMedian?: number
				readonly mcpHydratedOpenCountMedian?: number
				readonly mcpWarmHydratedOpenCountMedian?: number
				readonly mcpDocumentCacheHitsMedian?: number
				readonly mcpWarmDocumentCacheHitsMedian?: number
				readonly mcpCompactDefaultPayloadBytesMedian?: number
				readonly mcpCompactDefaultCellsMedian?: number
				readonly tuiOpenCallsMedian?: number
				readonly tuiWarmOpenCallsMedian?: number
				readonly tuiHydratedOpenCountMedian?: number
				readonly tuiWarmHydratedOpenCountMedian?: number
				readonly tuiDocumentCacheHitsMedian?: number
				readonly tuiWarmDocumentCacheHitsMedian?: number
				readonly apiPartial?: boolean
				readonly apiCompactDefaultPartial?: boolean
				readonly cliPartial?: boolean
				readonly mcpPartial?: boolean
				readonly mcpCompactDefaultPartial?: boolean
				readonly tuiPartial?: boolean
				readonly mcpPayloadBytesMedian?: number
				readonly tuiFrameBytesMedian?: number
			}
		}
		expect(payload.summary?.fullOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.fullWarmOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.cappedOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.cappedWarmOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.apiFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.apiWarmFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.apiCompactDefaultMedianMs).toBeNumber()
		expect(payload.summary?.apiWarmCompactDefaultMedianMs).toBeNumber()
		expect(payload.summary?.cliReadFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.cliWarmReadFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.mcpFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.mcpWarmFirstWindowMedianMs).toBeNumber()
		expect(payload.summary?.mcpCompactDefaultMedianMs).toBeNumber()
		expect(payload.summary?.mcpWarmCompactDefaultMedianMs).toBeNumber()
		expect(payload.summary?.tuiFirstPaintMedianMs).toBeNumber()
		expect(payload.summary?.tuiWarmFirstPaintMedianMs).toBeNumber()
		expect(payload.summary?.tuiOpenMedianMs).toBeNumber()
		expect(payload.summary?.tuiWarmOpenMedianMs).toBeNumber()
		expect(payload.summary?.tuiRenderMedianMs).toBeNumber()
		expect(payload.summary?.tuiWarmRenderMedianMs).toBeNumber()
		expect(payload.summary?.tuiHydrateMedianMs).toBeNumber()
		expect(payload.summary?.tuiWarmHydrateMedianMs).toBeNumber()
		expect(payload.summary?.fullRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.fullRetainedRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.cappedRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.cappedRetainedRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.apiRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.apiRetainedRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.mcpRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.mcpRetainedRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.tuiRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.tuiRetainedRssDeltaMbMedian).toBeNumber()
		expect(payload.summary?.cellsMedian).toBe(500 * 8)
		expect(payload.summary?.payloadBytesMedian).toBeGreaterThan(0)
		expect(payload.summary?.apiCompactDefaultPayloadBytesMedian).toBeGreaterThan(0)
		expect(payload.summary?.apiCompactDefaultCellsMedian).toBe(500 * 8)
		expect(payload.summary?.fullHydratedCellsMedian).toBe(800 * 8)
		expect(payload.summary?.cappedHydratedCellsMedian).toBe(500 * 8)
		expect(payload.summary?.tuiHydratedCellsMedian).toBe(500 * 8)
		expect(payload.summary?.fullOpenCallsMedian).toBe(1)
		expect(payload.summary?.fullWarmOpenCallsMedian).toBe(1)
		expect(payload.summary?.fullHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.fullWarmHydratedOpenCountMedian).toBe(0)
		expect(payload.summary?.fullDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.fullWarmDocumentCacheHitsMedian).toBe(1)
		expect(payload.summary?.cappedOpenCallsMedian).toBe(1)
		expect(payload.summary?.cappedWarmOpenCallsMedian).toBe(1)
		expect(payload.summary?.cappedHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.cappedWarmHydratedOpenCountMedian).toBe(0)
		expect(payload.summary?.cappedDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.cappedWarmDocumentCacheHitsMedian).toBe(1)
		expect(payload.summary?.apiOpenCallsMedian).toBe(1)
		expect(payload.summary?.apiWarmOpenCallsMedian).toBe(1)
		expect(payload.summary?.apiHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.apiWarmHydratedOpenCountMedian).toBe(0)
		expect(payload.summary?.apiDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.apiWarmDocumentCacheHitsMedian).toBe(1)
		expect(payload.summary?.cliOpenCallsMedian).toBe(1)
		expect(payload.summary?.cliWarmOpenCallsMedian).toBe(1)
		expect(payload.summary?.cliHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.cliWarmHydratedOpenCountMedian).toBe(0)
		expect(payload.summary?.cliDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.cliWarmDocumentCacheHitsMedian).toBe(1)
		expect(payload.summary?.mcpOpenCallsMedian).toBe(1)
		expect(payload.summary?.mcpWarmOpenCallsMedian).toBe(1)
		expect(payload.summary?.mcpHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.mcpWarmHydratedOpenCountMedian).toBe(0)
		expect(payload.summary?.mcpDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.mcpWarmDocumentCacheHitsMedian).toBe(1)
		expect(payload.summary?.mcpCompactDefaultPayloadBytesMedian).toBeGreaterThan(0)
		expect(payload.summary?.mcpCompactDefaultCellsMedian).toBe(500 * 8)
		expect(payload.summary?.tuiOpenCallsMedian).toBe(1)
		expect(payload.summary?.tuiWarmOpenCallsMedian).toBe(1)
		expect(payload.summary?.tuiHydratedOpenCountMedian).toBe(1)
		expect(payload.summary?.tuiWarmHydratedOpenCountMedian).toBe(0)
		expect(payload.summary?.tuiDocumentCacheHitsMedian).toBe(0)
		expect(payload.summary?.tuiWarmDocumentCacheHitsMedian).toBe(1)
		expect(payload.summary?.apiPartial).toBe(true)
		expect(payload.summary?.apiCompactDefaultPartial).toBe(true)
		expect(payload.summary?.cliPartial).toBe(true)
		expect(payload.summary?.mcpPartial).toBe(true)
		expect(payload.summary?.mcpCompactDefaultPartial).toBe(true)
		expect(payload.summary?.tuiPartial).toBe(true)
		expect(payload.summary?.mcpPayloadBytesMedian).toBeGreaterThan(0)
		expect(payload.summary?.tuiFrameBytesMedian).toBeGreaterThan(0)
	})

	test('can isolate one first-window case for profiling', async () => {
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
				'--only',
				'capped',
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
			readonly args?: { readonly only?: string }
			readonly summary?: {
				readonly fullOpenWindowMedianMs?: number
				readonly cappedOpenWindowMedianMs?: number
				readonly cappedWarmOpenWindowMedianMs?: number
				readonly cappedHydratedCellsMedian?: number
				readonly apiFirstWindowMedianMs?: number
				readonly mcpFirstWindowMedianMs?: number
				readonly tuiFirstPaintMedianMs?: number
				readonly cellsMedian?: number
			}
		}
		expect(payload.args?.only).toBe('capped')
		expect(payload.summary?.cappedOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.cappedWarmOpenWindowMedianMs).toBeNumber()
		expect(payload.summary?.cappedHydratedCellsMedian).toBe(500 * 8)
		expect(payload.summary?.cellsMedian).toBe(500 * 8)
		expect(payload.summary?.fullOpenWindowMedianMs).toBeUndefined()
		expect(payload.summary?.apiFirstWindowMedianMs).toBeUndefined()
		expect(payload.summary?.mcpFirstWindowMedianMs).toBeUndefined()
		expect(payload.summary?.tuiFirstPaintMedianMs).toBeUndefined()
	})

	test('runs against an existing input workbook without deleting it', async () => {
		const proc = Bun.spawn(
			[
				Bun.argv[0],
				runnerPath,
				'--input-file',
				'fixtures/xlsx/poi/SampleSS.xlsx',
				'--range',
				'A1:D10',
				'--row-limit',
				'5',
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
				readonly cleanup?: boolean
				readonly source?: string
			}
			readonly summary?: {
				readonly cellsMedian?: number
				readonly apiPartial?: boolean
				readonly cliPartial?: boolean
				readonly mcpPartial?: boolean
				readonly tuiPartial?: boolean
				readonly fullRetainedRssDeltaMbMedian?: number
			}
		}
		expect(payload.input).toEqual({
			xlsxPath: 'fixtures/xlsx/poi/SampleSS.xlsx',
			range: 'A1:D10',
			sheet: 'First Sheet',
			cleanup: false,
			source: 'input-file',
		})
		expect(payload.summary?.cellsMedian).toBeGreaterThan(0)
		expect(payload.summary?.apiPartial).toBe(true)
		expect(payload.summary?.cliPartial).toBe(true)
		expect(payload.summary?.mcpPartial).toBe(true)
		expect(payload.summary?.tuiPartial).toBe(true)
		expect(payload.summary?.fullRetainedRssDeltaMbMedian).toBeNumber()
	}, 20_000)
})
