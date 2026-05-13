#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { createApiFetch } from '../../apps/api/src/server.ts'
import { createServer } from '../../apps/mcp/src/index.ts'
import { indexToColumn } from '../../packages/core/src/index.ts'
import { WorkbookDocument } from '../../packages/sdk/src/index.ts'
import { buildRawReadWorkloadDataSet, type WorkloadName } from './competitive-io.ts'

interface Args {
	readonly rows: number
	readonly cols: number
	readonly rowLimit: number
	readonly workload: WorkloadName
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

interface Sample {
	readonly fullOpenWindowMs?: number
	readonly cappedOpenWindowMs?: number
	readonly apiFirstWindowMs?: number
	readonly mcpFirstWindowMs?: number
	readonly cells: number
	readonly payloadBytes?: number
	readonly mcpPayloadBytes?: number
	readonly fullHydratedCells?: number | null
	readonly cappedHydratedCells?: number | null
	readonly apiPartial?: boolean
	readonly mcpPartial?: boolean
}

const WORKLOADS = new Set<string>([
	'dense-values',
	'mixed-10pct-text',
	'mixed-50pct-text',
	'plain-text',
	'string-heavy',
	'sparse-wide',
])

function readOption(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function hasFlag(args: readonly string[], name: string): boolean {
	return args.includes(name)
}

function positiveInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value > 0 ? value : fallback
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

function parseArgs(): Args {
	const workload = readOption(process.argv, '--workload') ?? 'mixed-10pct-text'
	if (!WORKLOADS.has(workload)) throw new Error(`Unsupported --workload "${workload}"`)
	return {
		rows: positiveInt(readOption(process.argv, '--rows'), 65_536),
		cols: positiveInt(readOption(process.argv, '--cols'), 10),
		rowLimit: positiveInt(readOption(process.argv, '--row-limit'), 500),
		workload: workload as WorkloadName,
		repeat: positiveInt(readOption(process.argv, '--repeat'), 5),
		warmup: nonNegativeInt(readOption(process.argv, '--warmup'), 1),
		json: hasFlag(process.argv, '--json'),
	}
}

function median(values: readonly number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const middle = Math.floor(sorted.length / 2)
	const upper = sorted[middle] ?? 0
	return sorted.length % 2 === 1 ? upper : ((sorted[middle - 1] ?? upper) + upper) / 2
}

function medianOptional(values: readonly (number | undefined | null)[]): number | undefined {
	const defined = values.filter((value): value is number => typeof value === 'number')
	return defined.length > 0 ? median(defined) : undefined
}

async function time<T>(fn: () => Promise<T>): Promise<{ readonly ms: number; readonly result: T }> {
	const start = performance.now()
	const result = await fn()
	return { ms: performance.now() - start, result }
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

async function runFullOpenWindow(
	path: string,
	range: string,
	rowLimit: number,
): Promise<Pick<Sample, 'fullOpenWindowMs' | 'cells' | 'fullHydratedCells'>> {
	WorkbookDocument.clearCache()
	const measured = await time(async () => {
		const document = await WorkbookDocument.open(path, { mode: 'values' })
		const info = document.inspect()
		const window = document.readWindowCompact(info.sheets[0]?.name ?? 'Sheet1', range, {
			rowLimit,
			includeRefs: false,
			omitEmpty: true,
			flatValues: true,
		})
		return { info, window }
	})
	return {
		fullOpenWindowMs: measured.ms,
		cells: measured.result.window?.cells.length ?? 0,
		fullHydratedCells: measured.result.info.cellCount,
	}
}

async function runCappedOpenWindow(
	path: string,
	range: string,
	rowLimit: number,
): Promise<Pick<Sample, 'cappedOpenWindowMs' | 'cells' | 'cappedHydratedCells'>> {
	WorkbookDocument.clearCache()
	const measured = await time(async () => {
		const document = await WorkbookDocument.open(path, { mode: 'values', maxRows: rowLimit })
		const info = document.inspect()
		const window = document.readWindowCompact(info.sheets[0]?.name ?? 'Sheet1', range, {
			rowLimit,
			includeRefs: false,
			omitEmpty: true,
			flatValues: true,
		})
		return { info, window }
	})
	return {
		cappedOpenWindowMs: measured.ms,
		cells: measured.result.window?.cells.length ?? 0,
		cappedHydratedCells: measured.result.info.cellCount,
	}
}

async function runApiFirstWindow(
	path: string,
	range: string,
	rowLimit: number,
): Promise<Pick<Sample, 'apiFirstWindowMs' | 'cells' | 'payloadBytes' | 'apiPartial'>> {
	WorkbookDocument.clearCache()
	const apiFetch = createApiFetch()
	const body = JSON.stringify({
		file: path,
		range,
		format: 'compact',
		rowLimit,
	})
	const measured = await time(async () => {
		const response = await apiFetch(
			new Request('http://ascend.local/read', {
				method: 'POST',
				headers: { 'Content-Type': 'application/json' },
				body,
			}),
		)
		const text = await response.text()
		if (response.status !== 200) throw new Error(text)
		return { text, payload: JSON.parse(text) as ApiEnvelope }
	})
	const data = measured.result.payload.data
	return {
		apiFirstWindowMs: measured.ms,
		cells: data?.cells?.length ?? 0,
		payloadBytes: measured.result.text.length,
		apiPartial: data?.load?.isPartial ?? false,
	}
}

interface ApiEnvelope {
	readonly data?: {
		readonly cells?: readonly unknown[]
		readonly load?: { readonly isPartial?: boolean }
	}
}

interface McpReadResult {
	readonly structuredContent?: {
		readonly ok?: boolean
		readonly data?: {
			readonly cells?: readonly unknown[]
			readonly load?: { readonly isPartial?: boolean }
		}
		readonly error?: unknown
	}
}

type McpReadHandler = (args: {
	readonly file: string
	readonly range: string
	readonly format: 'compact'
	readonly rowLimit: number
}) => Promise<McpReadResult>

async function runMcpFirstWindow(
	path: string,
	range: string,
	rowLimit: number,
): Promise<Pick<Sample, 'mcpFirstWindowMs' | 'cells' | 'mcpPayloadBytes' | 'mcpPartial'>> {
	WorkbookDocument.clearCache()
	const server = createServer()
	const handler = (
		server as unknown as { _registeredTools: Record<string, { handler: McpReadHandler }> }
	)._registeredTools['ascend.read']?.handler
	if (!handler) throw new Error('MCP ascend.read handler not registered')
	const measured = await time(() =>
		handler({
			file: path,
			range,
			format: 'compact',
			rowLimit,
		}),
	)
	const content = measured.result.structuredContent
	if (content?.ok !== true) {
		throw new Error(`MCP ascend.read failed: ${JSON.stringify(content?.error ?? content)}`)
	}
	const data = content.data
	return {
		mcpFirstWindowMs: measured.ms,
		cells: data?.cells?.length ?? 0,
		mcpPayloadBytes: JSON.stringify(content).length,
		mcpPartial: data?.load?.isPartial ?? false,
	}
}

function summarize(samples: readonly Sample[]) {
	const fullOpenWindowMedianMs = medianOptional(samples.map((sample) => sample.fullOpenWindowMs))
	const cappedOpenWindowMedianMs = medianOptional(
		samples.map((sample) => sample.cappedOpenWindowMs),
	)
	const apiFirstWindowMedianMs = medianOptional(samples.map((sample) => sample.apiFirstWindowMs))
	const mcpFirstWindowMedianMs = medianOptional(samples.map((sample) => sample.mcpFirstWindowMs))
	return {
		fullOpenWindowMedianMs,
		cappedOpenWindowMedianMs,
		apiFirstWindowMedianMs,
		mcpFirstWindowMedianMs,
		...(fullOpenWindowMedianMs !== undefined && cappedOpenWindowMedianMs !== undefined
			? { cappedSpeedupVsFull: fullOpenWindowMedianMs / cappedOpenWindowMedianMs }
			: {}),
		...(fullOpenWindowMedianMs !== undefined && apiFirstWindowMedianMs !== undefined
			? { apiSpeedupVsFull: fullOpenWindowMedianMs / apiFirstWindowMedianMs }
			: {}),
		...(fullOpenWindowMedianMs !== undefined && mcpFirstWindowMedianMs !== undefined
			? { mcpSpeedupVsFull: fullOpenWindowMedianMs / mcpFirstWindowMedianMs }
			: {}),
		cellsMedian: medianOptional(samples.map((sample) => sample.cells)),
		payloadBytesMedian: medianOptional(samples.map((sample) => sample.payloadBytes)),
		mcpPayloadBytesMedian: medianOptional(samples.map((sample) => sample.mcpPayloadBytes)),
		fullHydratedCellsMedian: medianOptional(samples.map((sample) => sample.fullHydratedCells)),
		cappedHydratedCellsMedian: medianOptional(samples.map((sample) => sample.cappedHydratedCells)),
		apiPartial: samples.some((sample) => sample.apiPartial === true),
		mcpPartial: samples.some((sample) => sample.mcpPartial === true),
	}
}

async function run() {
	const args = parseArgs()
	const data = await buildRawReadWorkloadDataSet(args.workload, args.rows, args.cols)
	const range = `A1:${indexToColumn(args.cols - 1)}${args.rows}`
	const samples: Sample[] = []
	try {
		for (let i = 0; i < args.warmup; i++) {
			await runFullOpenWindow(data.xlsxPath, range, args.rowLimit)
			await runCappedOpenWindow(data.xlsxPath, range, args.rowLimit)
			await runApiFirstWindow(data.xlsxPath, range, args.rowLimit)
			await runMcpFirstWindow(data.xlsxPath, range, args.rowLimit)
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			const full = await runFullOpenWindow(data.xlsxPath, range, args.rowLimit)
			runGc()
			const capped = await runCappedOpenWindow(data.xlsxPath, range, args.rowLimit)
			runGc()
			const api = await runApiFirstWindow(data.xlsxPath, range, args.rowLimit)
			runGc()
			const mcp = await runMcpFirstWindow(data.xlsxPath, range, args.rowLimit)
			runGc()
			samples.push({ ...full, ...capped, ...api, ...mcp, cells: mcp.cells })
		}
		const payload = {
			tool: 'agent-first-window',
			args,
			summary: summarize(samples),
			samples,
		}
		if (args.json) console.log(JSON.stringify(payload, null, 2))
		else console.log(payload.summary)
	} finally {
		await rm(data.xlsxPath, { force: true })
	}
}

await run()
