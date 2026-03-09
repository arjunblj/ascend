import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { AscendWorkbook } from '../../packages/sdk/src/index.ts'

interface TimingResult {
	readonly name: string
	readonly durationMs: number
}

function sha256(bytes: Uint8Array): string {
	return createHash('sha256').update(bytes).digest('hex')
}

async function time<T>(
	name: string,
	fn: () => Promise<T>,
): Promise<{ result: T; timing: TimingResult }> {
	const start = performance.now()
	const result = await fn()
	return {
		result,
		timing: {
			name,
			durationMs: performance.now() - start,
		},
	}
}

function renderTimings(timings: readonly TimingResult[]): string {
	const headers = ['step', 'ms']
	const rows = timings.map((timing) => [timing.name, timing.durationMs.toFixed(2)])
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
	const argPath = process.argv[2]
	const target = resolve(argPath ?? 'research/real-workbooks/Book1.xlsx')
	const originalBytes = new Uint8Array(await readFile(target))
	const originalSha = sha256(originalBytes)

	const timings: TimingResult[] = []

	const metadata = await time('open-metadata', () =>
		AscendWorkbook.open(target, { mode: 'metadata-only' }),
	)
	timings.push(metadata.timing)

	const full = await time('open-full', () => AscendWorkbook.open(target))
	timings.push(full.timing)

	const bytes = await time('no-op-save-bytes', async () => full.result.toBytes())
	timings.push(bytes.timing)

	const parity = {
		byteIdentical: originalSha === sha256(bytes.result),
		sha256Before: originalSha,
		sha256After: sha256(bytes.result),
	}

	const info = full.result.inspect()

	const output = {
		file: target,
		parity,
		workbook: {
			sheetCount: info.sheetCount,
			loadedSheetCount: info.loadedSheetCount,
			cellCount: info.cellCount,
			workbookViewCount: info.workbookViewCount,
			externalReferenceCount: info.externalReferenceCount,
			compatibility: info.compatibility.status,
			styleSummary: info.styleSummary,
		},
		timings,
	}

	if (process.argv.includes('--json')) {
		console.log(JSON.stringify(output, null, 2))
		return
	}

	console.log(`Real workbook benchmark: ${target}`)
	console.log(`Byte-identical no-op save: ${parity.byteIdentical ? 'yes' : 'no'}`)
	console.log(renderTimings(timings))
}

await main()
