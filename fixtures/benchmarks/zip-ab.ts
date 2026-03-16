import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { unzipSync } from 'fflate'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'

interface Sample {
	readonly name: string
	readonly customMs: number
	readonly fflateMs: number
	readonly speedup: number
}

const FIXTURE_FILES = [
	'../xlsx/stress/dense-100k.xlsx',
	'../xlsx/stress/many-strings.xlsx',
	'../xlsx/poi/ConditionalFormattingSamples.xlsx',
] as const

function main(): void {
	const repeat = Math.max(1, Number.parseInt(process.argv[2] ?? '8', 10) || 8)
	const baseDir = fileURLToPath(new URL('.', import.meta.url))
	const samples: Sample[] = []

	for (const relativePath of FIXTURE_FILES) {
		const filePath = join(baseDir, relativePath)
		const bytes = new Uint8Array(readFileSync(filePath))
		warmup(bytes)
		const customMs = time(() => {
			const archive = extractZip(bytes)
			for (const entry of archive.entries()) {
				archive.readBytes(entry.path)
			}
		}, repeat)
		const fflateMs = time(() => {
			unzipSync(bytes)
		}, repeat)
		samples.push({
			name: relativePath.replace('../xlsx/', ''),
			customMs,
			fflateMs,
			speedup: customMs / fflateMs,
		})
	}

	console.log('ZIP Extraction A/B')
	console.log('='.repeat(88))
	console.log(
		[
			'Fixture'.padEnd(42),
			'Current(ms)'.padStart(12),
			'fflate(ms)'.padStart(12),
			'Speedup'.padStart(10),
		].join(' '),
	)
	console.log('-'.repeat(88))
	for (const sample of samples) {
		console.log(
			[
				sample.name.padEnd(42),
				sample.customMs.toFixed(2).padStart(12),
				sample.fflateMs.toFixed(2).padStart(12),
				`${sample.speedup.toFixed(2)}x`.padStart(10),
			].join(' '),
		)
	}
}

function warmup(bytes: Uint8Array): void {
	const archive = extractZip(bytes)
	for (const entry of archive.entries()) {
		archive.readBytes(entry.path)
	}
	unzipSync(bytes)
}

function time(fn: () => void, repeat: number): number {
	const runs: number[] = []
	for (let i = 0; i < repeat; i++) {
		const start = performance.now()
		fn()
		runs.push(performance.now() - start)
	}
	runs.sort((a, b) => a - b)
	return runs[Math.floor(runs.length / 2)] ?? 0
}

main()
