#!/usr/bin/env bun
import { indexToColumn } from '../../../packages/core/src/index.ts'
import { readXlsx } from '../../../packages/io-xlsx/src/index.ts'
import type { CellValue } from '../../../packages/schema/src/index.ts'

type Mode = 'values' | 'formula' | 'full' | 'metadata-only'
type Source = 'path' | 'bytes'

interface Args {
	readonly file: string
	readonly mode: Mode
	readonly source: Source
	readonly parseDates: boolean
	readonly materializeCells: boolean
	readonly repeat: number
	readonly warmup: number
	readonly json: boolean
}

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
	const argv = process.argv.slice(2)
	const file = readOption(argv, '--file')
	const mode = readOption(argv, '--mode') ?? 'values'
	const source = readOption(argv, '--source') ?? 'path'
	if (!file) throw new Error('--file is required')
	if (mode !== 'values' && mode !== 'formula' && mode !== 'full' && mode !== 'metadata-only') {
		throw new Error('--mode must be values, formula, full, or metadata-only')
	}
	if (source !== 'path' && source !== 'bytes') throw new Error('--source must be path or bytes')
	return {
		file,
		mode,
		source,
		parseDates: !hasFlag(argv, '--raw-values'),
		materializeCells: hasFlag(argv, '--materialize-cells'),
		repeat: positiveInt(readOption(argv, '--repeat'), 1),
		warmup: nonNegativeInt(readOption(argv, '--warmup'), 0),
		json: hasFlag(argv, '--json'),
	}
}

async function readBytes(args: Args, preloaded: Uint8Array | undefined): Promise<Uint8Array> {
	return preloaded ?? Bun.file(args.file).bytes()
}

async function openWorkbook(args: Args, preloaded: Uint8Array | undefined) {
	const bytes = await readBytes(args, preloaded)
	const result = readXlsx(bytes, {
		mode: args.mode,
		...(args.parseDates ? {} : { parseDates: false }),
	})
	if (!result.ok) throw new Error(result.error.message)
	return result.value
}

function memorySample(durationMs: number) {
	const memory = process.memoryUsage()
	const rss = typeof memory.rss === 'function' ? memory.rss() : memory.rss
	return {
		durationMs,
		rssAfterBytes: rss,
		peakRssBytes: rss,
		heapUsedBytes: memory.heapUsed,
		heapTotalBytes: memory.heapTotal,
		throughputPerSec: undefined,
	}
}

async function runMeasuredSample(args: Args): Promise<ReturnType<typeof memorySample>> {
	const proc = Bun.spawn(
		[
			'bun',
			'fixtures/benchmarks/runners/ascend_readxlsx_sample_worker.ts',
			'--file',
			args.file,
			'--mode',
			args.mode,
			'--source',
			args.source,
			...(args.warmup > 0 ? ['--worker-warmup', String(args.warmup)] : []),
			...(args.parseDates ? [] : ['--raw-values']),
			...(args.materializeCells ? ['--materialize-cells'] : []),
		],
		{ stdout: 'pipe', stderr: 'pipe' },
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) throw new Error(stderr || `sample worker exited ${exitCode}`)
	return JSON.parse(stdout) as ReturnType<typeof memorySample>
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

class OrderedHasher {
	private readonly hash = new Bun.CryptoHasher('sha256')
	private buffer = ''

	update(line: string): void {
		this.buffer += `${line.length}:${line}\n`
		if (this.buffer.length >= 1_048_576) this.flush()
	}

	digest(): string {
		this.flush()
		return this.hash.digest('hex')
	}

	private flush(): void {
		if (this.buffer.length === 0) return
		this.hash.update(this.buffer)
		this.buffer = ''
	}
}

function hashSorted(lines: readonly string[]): string {
	const hasher = new OrderedHasher()
	for (const line of [...lines].sort()) hasher.update(line)
	return hasher.digest()
}

function canonicalNumber(value: number): string {
	return Object.is(value, -0) ? '0' : String(value)
}

function serializeCellValue(value: CellValue): string {
	switch (value.kind) {
		case 'empty':
			return 'empty'
		case 'number':
			return `n:${canonicalNumber(value.value)}`
		case 'date':
			return `n:${canonicalNumber(value.serial)}`
		case 'string':
			return `s:${value.value}`
		case 'richText':
			return `s:${value.runs.map((run) => run.text).join('')}`
		case 'boolean':
			return `b:${value.value ? 'true' : 'false'}`
		case 'error':
			return `e:${value.value}`
	}
}

function readAssertions(
	result: Awaited<ReturnType<typeof openWorkbook>>,
	args: Args,
): Record<string, string | number | boolean | null> {
	const sheetNames = result.workbook.sheets.map((sheet) => sheet.name)
	let cellCount = 0
	let formulaCount = 0
	const usedRanges: string[] = []
	const orderedRefs = new OrderedHasher()
	const orderedValues = new OrderedHasher()
	const orderedFormulas = new OrderedHasher()
	for (const sheet of result.workbook.sheets) {
		formulaCount += sheet.cells.formulaCellCount()
		const usedRange = sheet.cells.usedRange()
		usedRanges.push(
			usedRange
				? `${sheet.name}!${indexToColumn(usedRange.start.col)}${usedRange.start.row + 1}:${indexToColumn(
						usedRange.end.col,
					)}${usedRange.end.row + 1}`
				: `${sheet.name}!empty`,
		)
		for (const [row, col, cell] of sheet.cells.iterate()) {
			cellCount++
			const ref = `${sheet.name}!${indexToColumn(col)}${row + 1}`
			orderedRefs.update(ref)
			orderedValues.update(`${ref}\t${serializeCellValue(cell.value)}`)
			if (cell.formula) orderedFormulas.update(`${ref}=${cell.formula}`)
		}
	}
	return {
		sheetCount: result.workbook.sheets.length,
		sheetNamesHash: hashSorted(sheetNames.map((name, index) => `${index}:${name}`)),
		cellCount,
		physicalCellCount: null,
		formulaCount,
		usedRangeCount: usedRanges.length,
		firstUsedRange: usedRanges[0] ?? null,
		firstPhysicalUsedRange: null,
		usedRangesHash: hashSorted(usedRanges),
		physicalUsedRangesHash: hashSorted([]),
		orderedSemanticCellRefsHash: orderedRefs.digest(),
		orderedSemanticCellValuesHash: orderedValues.digest(),
		orderedFormulaTextHash: orderedFormulas.digest(),
		runnerVersion: 'workspace',
		runnerSource: args.source,
		runnerLoadMode: args.mode,
		runnerRichMetadata: false,
		runnerParseDates: args.parseDates,
		runnerApi: 'readXlsx',
		runnerTimeOpenOnly: true,
		runnerAssertionMode: 'ordered-hashes-open-runner',
	}
}

const args = parseArgs()
for (let i = 0; i < args.warmup; i++) {
	await runMeasuredSample(args)
	runGc()
}
const samples: ReturnType<typeof memorySample>[] = []
for (let i = 0; i < args.repeat; i++) {
	runGc()
	samples.push(await runMeasuredSample(args))
}
runGc()
const preloaded = args.source === 'bytes' ? await Bun.file(args.file).bytes() : undefined
const finalWorkbook = await openWorkbook(args, preloaded)
const assertions = readAssertions(finalWorkbook, args)
const payload = { assertions, samples }
console.log(args.json ? JSON.stringify(payload) : JSON.stringify(payload, null, 2))
