#!/usr/bin/env bun
import { readXlsx } from '../../../packages/io-xlsx/src/index.ts'

type Mode = 'values' | 'formula' | 'full' | 'metadata-only'
type Source = 'path' | 'bytes'

function readOption(args: readonly string[], name: string): string | undefined {
	const index = args.indexOf(name)
	return index >= 0 ? args[index + 1] : undefined
}

function hasFlag(args: readonly string[], name: string): boolean {
	return args.includes(name)
}

function nonNegativeInt(raw: string | undefined, fallback: number): number {
	const value = raw ? Number.parseInt(raw, 10) : fallback
	return Number.isFinite(value) && value >= 0 ? value : fallback
}

function runGc(): void {
	;(Bun as unknown as { gc?: (force?: boolean) => void }).gc?.(true)
}

function memorySample() {
	const memory = process.memoryUsage()
	const rss = typeof memory.rss === 'function' ? memory.rss() : memory.rss
	return { memory, rss }
}

const argv = process.argv.slice(2)
const file = readOption(argv, '--file')
const mode = (readOption(argv, '--mode') ?? 'values') as Mode
const source = (readOption(argv, '--source') ?? 'path') as Source
const workerWarmup = nonNegativeInt(readOption(argv, '--worker-warmup'), 0)
if (!file) throw new Error('--file is required')
if (source !== 'path' && source !== 'bytes') throw new Error('--source must be path or bytes')
const preloadedBytes = source === 'bytes' ? await Bun.file(file).bytes() : undefined
for (let i = 0; i < workerWarmup; i++) {
	const bytes = preloadedBytes ?? (await Bun.file(file).bytes())
	const result = readXlsx(bytes, {
		mode,
		...(hasFlag(argv, '--raw-values') ? { parseDates: false } : {}),
	})
	if (!result.ok) throw new Error(result.error.message)
}
runGc()
const before = memorySample()
const start = performance.now()
const bytes = preloadedBytes ?? (await Bun.file(file).bytes())
const result = readXlsx(bytes, {
	mode,
	...(hasFlag(argv, '--raw-values') ? { parseDates: false } : {}),
})
if (!result.ok) throw new Error(result.error.message)
if (hasFlag(argv, '--materialize-cells')) {
	let count = 0
	for (const sheet of result.value.workbook.sheets) {
		for (const _cell of sheet.cells.iterate()) count++
	}
	if (count < 0) throw new Error('unreachable')
}
const durationMs = performance.now() - start
const after = memorySample()
runGc()
const afterGc = memorySample()
console.log(
	JSON.stringify({
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
	}),
)
