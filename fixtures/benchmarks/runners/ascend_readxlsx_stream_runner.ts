#!/usr/bin/env bun
import { indexToColumn } from '../../../packages/core/src/index.ts'
import { readXlsxRowsStream } from '../../../packages/io-xlsx/src/index.ts'
import type { CellValue } from '../../../packages/schema/src/index.ts'

type Mode = 'values' | 'formula'
type Source = 'path' | 'bytes'

interface Args {
	readonly file: string
	readonly mode: Mode
	readonly source: Source
	readonly chunkedSheetXml: boolean
	readonly sampleWorker: boolean
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
	const source = readOption(argv, '--source') ?? 'bytes'
	if (!file) throw new Error('--file is required')
	if (mode !== 'values' && mode !== 'formula') throw new Error('--mode must be values or formula')
	if (source !== 'path' && source !== 'bytes') throw new Error('--source must be path or bytes')
	return {
		file,
		mode,
		source,
		chunkedSheetXml: hasFlag(argv, '--chunked-sheet-xml'),
		sampleWorker: hasFlag(argv, '--sample-worker'),
		repeat: positiveInt(readOption(argv, '--repeat'), 1),
		warmup: nonNegativeInt(readOption(argv, '--warmup'), 0),
		json: hasFlag(argv, '--json'),
	}
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function memorySnapshot() {
	const memory = process.memoryUsage()
	const rss = typeof memory.rss === 'function' ? memory.rss() : memory.rss
	return { memory, rss }
}

function memorySample(durationMs: number, before = memorySnapshot()) {
	const after = memorySnapshot()
	runGc()
	const afterGc = memorySnapshot()
	return {
		durationMs,
		rssDeltaBytes: Math.max(0, after.rss - before.rss),
		retainedRssDeltaBytes: Math.max(0, afterGc.rss - before.rss),
		rssAfterBytes: after.rss,
		rssAfterGcBytes: afterGc.rss,
		peakRssBytes: Math.max(before.rss, after.rss, afterGc.rss),
		heapDeltaBytes: Math.max(0, after.memory.heapUsed - before.memory.heapUsed),
		heapUsedBytes: after.memory.heapUsed,
		heapTotalBytes: after.memory.heapTotal,
		heapAfterGcBytes: afterGc.memory.heapUsed,
	}
}

async function readSource(
	args: Args,
	preloaded: Uint8Array | undefined,
): Promise<Uint8Array | string> {
	return args.source === 'bytes' ? (preloaded ?? Bun.file(args.file).bytes()) : args.file
}

async function consumeStream(args: Args, preloaded: Uint8Array | undefined): Promise<number> {
	const source = await readSource(args, preloaded)
	const result = await readXlsxRowsStream(
		typeof source === 'string' ? Bun.file(source).stream() : source,
		{
			mode: args.mode,
			...(args.chunkedSheetXml ? { chunkedSheetXml: true } : {}),
		},
	)
	if (!result.ok) throw new Error(result.error.message)
	let count = 0
	for await (const row of result.value) count += row.cells.length
	return count
}

async function runSampleWorker(args: Args): Promise<void> {
	const preloaded = args.source === 'bytes' ? await Bun.file(args.file).bytes() : undefined
	for (let i = 0; i < args.warmup; i++) await consumeStream(args, preloaded)
	runGc()
	const before = memorySnapshot()
	const start = performance.now()
	const count = await consumeStream(args, preloaded)
	const sample = {
		...memorySample(performance.now() - start, before),
		streamedCellCount: count,
	}
	console.log(JSON.stringify(sample))
}

async function runMeasuredSample(args: Args): Promise<Record<string, number>> {
	const proc = Bun.spawn(
		[
			'bun',
			'fixtures/benchmarks/runners/ascend_readxlsx_stream_runner.ts',
			'--sample-worker',
			'--file',
			args.file,
			'--mode',
			args.mode,
			'--source',
			args.source,
			...(args.chunkedSheetXml ? ['--chunked-sheet-xml'] : []),
			...(args.warmup > 0 ? ['--warmup', String(args.warmup)] : []),
		],
		{ stdout: 'pipe', stderr: 'pipe' },
	)
	const [stdout, stderr, exitCode] = await Promise.all([
		new Response(proc.stdout).text(),
		new Response(proc.stderr).text(),
		proc.exited,
	])
	if (exitCode !== 0) throw new Error(stderr || `stream sample worker exited ${exitCode}`)
	return JSON.parse(stdout) as Record<string, number>
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

async function streamAssertions(
	args: Args,
): Promise<Record<string, string | number | boolean | null>> {
	const source = await readSource(
		args,
		args.source === 'bytes' ? await Bun.file(args.file).bytes() : undefined,
	)
	const result = await readXlsxRowsStream(
		typeof source === 'string' ? Bun.file(source).stream() : source,
		{
			mode: args.mode,
			...(args.chunkedSheetXml ? { chunkedSheetXml: true } : {}),
		},
	)
	if (!result.ok) throw new Error(result.error.message)
	const sheetName = 'Data'
	const orderedRefs = new OrderedHasher()
	const orderedValues = new OrderedHasher()
	const orderedFormulas = new OrderedHasher()
	let cellCount = 0
	let firstRow: number | undefined
	let lastRow: number | undefined
	let maxCol = -1
	for await (const row of result.value) {
		if (row.cells.length === 0) continue
		firstRow ??= row.row
		lastRow = row.row
		for (const [col, cell] of row.cells) {
			cellCount++
			if (col > maxCol) maxCol = col
			const ref = `${sheetName}!${indexToColumn(col)}${row.row + 1}`
			orderedRefs.update(ref)
			orderedValues.update(`${ref}\t${serializeCellValue(cell.value)}`)
			if (cell.formula) orderedFormulas.update(`${ref}=${cell.formula}`)
		}
	}
	const usedRange =
		cellCount > 0 && firstRow !== undefined && lastRow !== undefined && maxCol >= 0
			? `${sheetName}!A${firstRow + 1}:${indexToColumn(maxCol)}${lastRow + 1}`
			: `${sheetName}!empty`
	return {
		sheetCount: 1,
		sheetNamesHash: hashSorted([`0:${sheetName}`]),
		cellCount,
		physicalCellCount: cellCount,
		formulaCount: 0,
		usedRangeCount: 1,
		firstUsedRange: usedRange,
		firstPhysicalUsedRange: usedRange,
		usedRangesHash: hashSorted([usedRange]),
		physicalUsedRangesHash: hashSorted([usedRange]),
		orderedSemanticCellRefsHash: orderedRefs.digest(),
		orderedSemanticCellValuesHash: orderedValues.digest(),
		orderedFormulaTextHash: orderedFormulas.digest(),
		runnerVersion: 'workspace',
		runnerSource: args.source,
		runnerLoadMode: args.mode,
		runnerApi: 'readXlsxRowsStream',
		runnerChunkedSheetXml: args.chunkedSheetXml,
		runnerAssertionMode: 'ordered-hashes-stream-runner',
	}
}

const args = parseArgs()
if (args.sampleWorker) {
	await runSampleWorker(args)
} else {
	const samples: Record<string, number>[] = []
	for (let i = 0; i < args.repeat; i++) samples.push(await runMeasuredSample(args))
	const assertions = await streamAssertions(args)
	const payload = { assertions, samples }
	console.log(args.json ? JSON.stringify(payload) : JSON.stringify(payload, null, 2))
}
