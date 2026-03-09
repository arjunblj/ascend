import { createHash } from 'node:crypto'
import { access, readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'
import {
	type BenchmarkCaseResult,
	createBenchmarkSuite,
	formatBytes,
	summarizeSamples,
} from './results.ts'

interface TimingResult {
	readonly name: string
	readonly durationMs: number
	readonly rssDeltaBytes?: number
	readonly rssAfterBytes?: number
	readonly retainedRssDeltaBytes?: number
	readonly rssAfterGcBytes?: number
	readonly samples?: readonly {
		readonly durationMs: number
		readonly rssDeltaBytes?: number
		readonly retainedRssDeltaBytes?: number
		readonly rssAfterBytes?: number
		readonly rssAfterGcBytes?: number
	}[]
}

interface StepResult {
	readonly timing: TimingResult
	readonly parity?: {
		readonly byteIdentical: boolean
		readonly sha256Before: string
		readonly sha256After: string
	}
	readonly workbook?: {
		readonly sheetCount: number
		readonly loadedSheetCount: number
		readonly cellCount: number | null
		readonly workbookViewCount: number
		readonly externalReferenceCount: number
		readonly compatibility: string
		readonly styleSummary: {
			readonly numFmtCount: number
			readonly fontCount: number
			readonly fillCount: number
			readonly borderCount: number
			readonly cellXfCount: number
			readonly dxfCount: number
			readonly tableStyleCount: number
		}
	}
	readonly assertions?: Record<string, string | number | boolean | null>
}

type StepName =
	| 'open-metadata'
	| 'open-values'
	| 'open-full'
	| 'read-window-values'
	| 'workflow-inspect-read'
	| 'preview-numeric-edit'
	| 'preview-format-edit'
	| 'no-op-save-bytes'
	| 'numeric-edit-save-bytes'
	| 'format-edit-save-bytes'

function sha256(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

async function time<T>(
	name: string,
	fn: () => Promise<T>,
): Promise<{ result: T; timing: TimingResult }> {
	runGc()
	const rssBefore = getRssBytes()
	const start = performance.now()
	const result = await fn()
	const rssAfter = getRssBytes()
	runGc()
	const rssAfterGc = getRssBytes()
	return {
		result,
		timing: {
			name,
			durationMs: performance.now() - start,
			rssDeltaBytes:
				rssBefore !== undefined && rssAfter !== undefined
					? Math.max(0, rssAfter - rssBefore)
					: undefined,
			rssAfterBytes: rssAfter,
			retainedRssDeltaBytes:
				rssBefore !== undefined && rssAfterGc !== undefined
					? Math.max(0, rssAfterGc - rssBefore)
					: undefined,
			rssAfterGcBytes: rssAfterGc,
		},
	}
}

function renderTimings(results: readonly BenchmarkCaseResult[]): string {
	const headers = [
		'step',
		'median-ms',
		'p95-ms',
		'rss-delta',
		'rss-after',
		'retained',
		'rss-after-gc',
	]
	const rows = results.map((result) => [
		result.name,
		result.metrics.medianMs.toFixed(2),
		result.metrics.p95Ms.toFixed(2),
		result.metrics.rssDeltaBytes !== undefined ? formatBytes(result.metrics.rssDeltaBytes) : 'n/a',
		result.metrics.rssAfterBytes !== undefined ? formatBytes(result.metrics.rssAfterBytes) : 'n/a',
		result.metrics.retainedRssDeltaBytes !== undefined
			? formatBytes(result.metrics.retainedRssDeltaBytes)
			: 'n/a',
		result.metrics.rssAfterGcBytes !== undefined
			? formatBytes(result.metrics.rssAfterGcBytes)
			: 'n/a',
	])
	const widths = headers.map((header, index) =>
		Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)),
	)
	const pad = (value: string, width: number) =>
		value + ' '.repeat(Math.max(0, width - value.length))
	const line = (cells: readonly string[]) =>
		cells.map((cell, index) => pad(cell, widths[index] ?? 0)).join('  ')
	return [
		line(headers),
		widths.map((width) => '─'.repeat(width)).join('──'),
		...rows.map(line),
	].join('\n')
}

async function main(): Promise<void> {
	const argPath = readPositionalArg(0)
	const target = argPath ? resolve(argPath) : await defaultBenchmarkTarget()
	const step = readFlag('--step')
	const repeat = Math.max(1, Number.parseInt(readFlag('--repeat') ?? '1', 10) || 1)
	if (step) {
		const result = await runStep(target, step as StepName)
		console.log(JSON.stringify(result, null, 2))
		return
	}

	const stepNames: readonly StepName[] = [
		'open-metadata',
		'open-values',
		'open-full',
		'read-window-values',
		'workflow-inspect-read',
		'preview-numeric-edit',
		'preview-format-edit',
		'no-op-save-bytes',
		'numeric-edit-save-bytes',
		'format-edit-save-bytes',
	]
	const results: StepResult[] = []
	for (const name of stepNames) {
		results.push(await runRepeatedStep(target, name, repeat))
	}
	const cases = results.map((result) => toBenchmarkCase(result))
	const parity = results.find((result) => result.parity)?.parity
	const workbook = results.find((result) => result.workbook)?.workbook

	const output = createBenchmarkSuite({
		suite: 'ascend-real-workbook-benchmarks',
		kind: 'real-workbook',
		cases,
		metadata: {
			file: target,
			repeat,
			...(parity ? { parity } : {}),
			...(workbook ? { workbook } : {}),
		},
	})

	if (process.argv.includes('--json')) {
		console.log(JSON.stringify(output, null, 2))
		return
	}

	console.log(`Real workbook benchmark: ${target}`)
	console.log(`Byte-identical no-op save: ${parity?.byteIdentical ? 'yes' : 'no'}`)
	console.log(renderTimings(cases))
}

async function runIsolatedStep(target: string, step: StepName): Promise<StepResult> {
	const proc = Bun.spawn(
		['bun', 'run', process.argv[1] ?? import.meta.path, target, '--step', step],
		{
			stdout: 'pipe',
			stderr: 'pipe',
			cwd: process.cwd(),
		},
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) {
		throw new Error(stderr.trim() || `Benchmark step "${step}" failed`)
	}
	return JSON.parse(stdout) as StepResult
}

async function runStep(target: string, step: StepName): Promise<StepResult> {
	const originalBytes = new Uint8Array(await readFile(target))
	const originalSha = sha256(originalBytes)

	switch (step) {
		case 'open-metadata': {
			const { timing } = await time('open-metadata', () =>
				AscendWorkbook.open(target, { mode: 'metadata-only' }),
			)
			return { timing }
		}
		case 'open-values': {
			const { timing } = await time('open-values', () =>
				AscendWorkbook.open(target, { mode: 'values' }),
			)
			return { timing }
		}
		case 'open-full': {
			const { result, timing } = await time('open-full', () => AscendWorkbook.open(target))
			const info = result.inspect()
			return {
				timing,
				workbook: {
					sheetCount: info.sheetCount,
					loadedSheetCount: info.loadedSheetCount,
					cellCount: info.cellCount,
					workbookViewCount: info.workbookViewCount,
					externalReferenceCount: info.externalReferenceCount,
					compatibility: info.compatibility.status,
					styleSummary: info.styleSummary,
				},
			}
		}
		case 'read-window-values': {
			const wb = await AscendWorkbook.open(target, { mode: 'values' })
			const probe = pickReadProbe(wb)
			const { result, timing } = await time('read-window-values', async () =>
				Promise.resolve(wb.readWindow(probe.sheet, probe.range, { rowLimit: probe.rowLimit })),
			)
			if (!result) throw new Error('Read-window benchmark failed to load target range')
			return {
				timing,
				assertions: {
					sheet: probe.sheet,
					range: probe.range,
					returnedCells: result.cells.length,
					hasMore: result.hasMore,
				},
			}
		}
		case 'workflow-inspect-read': {
			const wb = await AscendWorkbook.open(target, { mode: 'values' })
			const probe = pickReadProbe(wb)
			const { result, timing } = await time('workflow-inspect-read', async () => {
				const info = wb.inspect()
				const first = wb.readWindow(probe.sheet, probe.range, { rowLimit: probe.rowLimit })
				const second = wb.readWindow(probe.sheet, probe.range, {
					rowOffset: probe.rowLimit,
					rowLimit: probe.rowLimit,
				})
				return {
					sheetCount: info.sheetCount,
					firstCells: first?.cells.length ?? 0,
					secondCells: second?.cells.length ?? 0,
				}
			})
			return {
				timing,
				assertions: {
					sheetCount: result.sheetCount,
					firstCells: result.firstCells,
					secondCells: result.secondCells,
				},
			}
		}
		case 'preview-numeric-edit': {
			const wb = await AscendWorkbook.open(target)
			const probe = pickNumericProbe(wb)
			const { timing } = await time('preview-numeric-edit', async () =>
				wb.preview([
					{ op: 'setCells', sheet: probe.sheet, updates: [{ ref: probe.ref, value: probe.value }] },
				]),
			)
			return { timing }
		}
		case 'preview-format-edit': {
			const wb = await AscendWorkbook.open(target)
			const probe = pickNumericProbe(wb)
			const { timing } = await time('preview-format-edit', async () =>
				wb.preview([
					{ op: 'setNumberFormat', sheet: probe.sheet, range: probe.ref, format: '0.0%' },
				]),
			)
			return { timing }
		}
		case 'no-op-save-bytes': {
			const wb = await AscendWorkbook.open(target)
			const { result, timing } = await time('no-op-save-bytes', async () => wb.toBytes())
			return {
				timing,
				parity: {
					byteIdentical: originalSha === sha256(result),
					sha256Before: originalSha,
					sha256After: sha256(result),
				},
			}
		}
		case 'numeric-edit-save-bytes': {
			const wb = await AscendWorkbook.open(target)
			const probe = pickNumericProbe(wb)
			wb.apply([
				{ op: 'setCells', sheet: probe.sheet, updates: [{ ref: probe.ref, value: probe.value }] },
			])
			const { result, timing } = await time('numeric-edit-save-bytes', async () => wb.toBytes())
			return {
				timing,
				parity: {
					byteIdentical: originalSha === sha256(result),
					sha256Before: originalSha,
					sha256After: sha256(result),
				},
			}
		}
		case 'format-edit-save-bytes': {
			const wb = await AscendWorkbook.open(target)
			const probe = pickNumericProbe(wb)
			wb.apply([{ op: 'setNumberFormat', sheet: probe.sheet, range: probe.ref, format: '0.0%' }])
			const { timing } = await time('format-edit-save-bytes', async () => wb.toBytes())
			return { timing }
		}
	}
}

async function runRepeatedStep(
	target: string,
	step: StepName,
	repeat: number,
): Promise<StepResult> {
	const results: StepResult[] = []
	for (let i = 0; i < repeat; i++) {
		results.push(await runIsolatedStep(target, step))
	}
	if (results.length === 1) return results[0] as StepResult

	const timings = results.map((result) => result.timing)
	const aggregateTiming: TimingResult = {
		name: step,
		durationMs: median(timings.map((timing) => timing.durationMs)),
		rssDeltaBytes: medianDefined(timings.map((timing) => timing.rssDeltaBytes)),
		rssAfterBytes: medianDefined(timings.map((timing) => timing.rssAfterBytes)),
		retainedRssDeltaBytes: medianDefined(timings.map((timing) => timing.retainedRssDeltaBytes)),
		rssAfterGcBytes: medianDefined(timings.map((timing) => timing.rssAfterGcBytes)),
		samples: timings.map((timing) => ({
			durationMs: timing.durationMs,
			...(timing.rssDeltaBytes !== undefined ? { rssDeltaBytes: timing.rssDeltaBytes } : {}),
			...(timing.retainedRssDeltaBytes !== undefined
				? { retainedRssDeltaBytes: timing.retainedRssDeltaBytes }
				: {}),
			...(timing.rssAfterBytes !== undefined ? { rssAfterBytes: timing.rssAfterBytes } : {}),
			...(timing.rssAfterGcBytes !== undefined ? { rssAfterGcBytes: timing.rssAfterGcBytes } : {}),
		})),
	}

	return {
		timing: aggregateTiming,
		parity: results.find((result) => result.parity)?.parity,
		workbook: results.find((result) => result.workbook)?.workbook,
		assertions: results.find((result) => result.assertions)?.assertions,
	}
}

function toBenchmarkCase(result: StepResult): BenchmarkCaseResult {
	const samples = result.timing.samples?.map((sample) => ({
		durationMs: sample.durationMs,
		...(sample.rssDeltaBytes !== undefined ? { rssDeltaBytes: sample.rssDeltaBytes } : {}),
		...(sample.retainedRssDeltaBytes !== undefined
			? { retainedRssDeltaBytes: sample.retainedRssDeltaBytes }
			: {}),
		...(sample.rssAfterBytes !== undefined ? { rssAfterBytes: sample.rssAfterBytes } : {}),
		...(sample.rssAfterGcBytes !== undefined ? { rssAfterGcBytes: sample.rssAfterGcBytes } : {}),
	})) ?? [
		{
			durationMs: result.timing.durationMs,
			...(result.timing.rssDeltaBytes !== undefined
				? { rssDeltaBytes: result.timing.rssDeltaBytes }
				: {}),
			...(result.timing.retainedRssDeltaBytes !== undefined
				? { retainedRssDeltaBytes: result.timing.retainedRssDeltaBytes }
				: {}),
			...(result.timing.rssAfterBytes !== undefined
				? { rssAfterBytes: result.timing.rssAfterBytes }
				: {}),
			...(result.timing.rssAfterGcBytes !== undefined
				? { rssAfterGcBytes: result.timing.rssAfterGcBytes }
				: {}),
		},
	]
	return {
		name: result.timing.name,
		category: 'workflow',
		dimensions: {},
		metrics: summarizeSamples(samples),
		...(result.timing.samples ? { samples } : {}),
		...(result.assertions ? { assertions: result.assertions } : {}),
	}
}

function readFlag(name: string): string | undefined {
	const index = process.argv.indexOf(name)
	return index >= 0 ? process.argv[index + 1] : undefined
}

async function defaultBenchmarkTarget(): Promise<string> {
	const candidates = [
		'research/excel-corpus/ms-excel-formulas-and-pivot-tables.xlsx',
		'research/excel-corpus/conditional-formatting.xlsx',
		'research/real-workbooks/Book1.xlsx',
	]
	for (const candidate of candidates) {
		const resolved = resolve(candidate)
		try {
			await access(resolved)
			return resolved
		} catch {}
	}
	return resolve(candidates[candidates.length - 1] ?? 'research/real-workbooks/Book1.xlsx')
}

function readPositionalArg(index: number): string | undefined {
	const positional = process.argv.slice(2).filter((arg, argIndex, args) => {
		if (arg.startsWith('--')) return false
		const previous = args[argIndex - 1]
		if (previous && previous.startsWith('--')) return false
		return true
	})
	return positional[index]
}

function pickNumericProbe(wb: AscendWorkbook): { sheet: string; ref: string; value: number } {
	for (const sheetName of wb.sheets) {
		const sheet = wb.sheet(sheetName)
		if (!sheet) continue
		const used = sheet.usedRange()
		if (!used) continue
		const range = `${columnLabel(used.start.col)}${used.start.row + 1}:${columnLabel(used.end.col)}${used.end.row + 1}`
		for (const row of sheet.streamRange(range)) {
			for (const cell of row) {
				if (cell.value.kind === 'number') {
					return { sheet: sheetName, ref: cell.ref, value: cell.value.value + 1 }
				}
			}
		}
	}
	return { sheet: wb.sheets[0] ?? 'Sheet1', ref: 'A1', value: 1 }
}

function pickReadProbe(wb: AscendWorkbook): { sheet: string; range: string; rowLimit: number } {
	for (const sheetName of wb.sheets) {
		const sheet = wb.sheet(sheetName)
		if (!sheet) continue
		const used = sheet.usedRange()
		if (!used) continue
		const endCol = Math.min(used.end.col, 19)
		const endRow = Math.min(used.end.row, 999)
		const rowLimit = Math.max(1, Math.min(200, endRow - used.start.row + 1))
		return {
			sheet: sheetName,
			range: `${columnLabel(used.start.col)}${used.start.row + 1}:${columnLabel(endCol)}${endRow + 1}`,
			rowLimit,
		}
	}
	return { sheet: wb.sheets[0] ?? 'Sheet1', range: 'A1:A1', rowLimit: 1 }
}

function median(values: number[]): number {
	const sorted = [...values].sort((a, b) => a - b)
	const mid = Math.floor(sorted.length / 2)
	return sorted.length % 2 === 0
		? ((sorted[mid - 1] ?? 0) + (sorted[mid] ?? 0)) / 2
		: (sorted[mid] ?? 0)
}

function medianDefined(values: Array<number | undefined>): number | undefined {
	const defined = values.filter((value): value is number => value !== undefined)
	return defined.length > 0 ? median(defined) : undefined
}

function columnLabel(col: number): string {
	let n = col
	let label = ''
	while (n >= 0) {
		label = String.fromCharCode(65 + (n % 26)) + label
		n = Math.floor(n / 26) - 1
	}
	return label
}

function getRssBytes(): number | undefined {
	try {
		return process.memoryUsage.rss()
	} catch {
		return undefined
	}
}

function runGc(): void {
	try {
		;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
	} catch {
		// Best effort only.
	}
}

await main()
