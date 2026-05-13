#!/usr/bin/env bun
import { rm } from 'node:fs/promises'
import { createApiFetch } from '../../apps/api/src/server.ts'
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
	readonly cells: number
	readonly payloadBytes?: number
	readonly fullHydratedCells?: number | null
	readonly cappedHydratedCells?: number | null
	readonly apiPartial?: boolean
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

function summarize(samples: readonly Sample[]) {
	const fullOpenWindowMedianMs = medianOptional(samples.map((sample) => sample.fullOpenWindowMs))
	const cappedOpenWindowMedianMs = medianOptional(
		samples.map((sample) => sample.cappedOpenWindowMs),
	)
	const apiFirstWindowMedianMs = medianOptional(samples.map((sample) => sample.apiFirstWindowMs))
	return {
		fullOpenWindowMedianMs,
		cappedOpenWindowMedianMs,
		apiFirstWindowMedianMs,
		...(fullOpenWindowMedianMs !== undefined && cappedOpenWindowMedianMs !== undefined
			? { cappedSpeedupVsFull: fullOpenWindowMedianMs / cappedOpenWindowMedianMs }
			: {}),
		...(fullOpenWindowMedianMs !== undefined && apiFirstWindowMedianMs !== undefined
			? { apiSpeedupVsFull: fullOpenWindowMedianMs / apiFirstWindowMedianMs }
			: {}),
		cellsMedian: medianOptional(samples.map((sample) => sample.cells)),
		payloadBytesMedian: medianOptional(samples.map((sample) => sample.payloadBytes)),
		fullHydratedCellsMedian: medianOptional(samples.map((sample) => sample.fullHydratedCells)),
		cappedHydratedCellsMedian: medianOptional(samples.map((sample) => sample.cappedHydratedCells)),
		apiPartial: samples.some((sample) => sample.apiPartial === true),
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
			runGc()
		}
		for (let i = 0; i < args.repeat; i++) {
			const full = await runFullOpenWindow(data.xlsxPath, range, args.rowLimit)
			runGc()
			const capped = await runCappedOpenWindow(data.xlsxPath, range, args.rowLimit)
			runGc()
			const api = await runApiFirstWindow(data.xlsxPath, range, args.rowLimit)
			runGc()
			samples.push({ ...full, ...capped, ...api, cells: api.cells })
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
