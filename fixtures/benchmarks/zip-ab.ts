import { readFileSync } from 'node:fs'
import { isAbsolute, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { deflateRawSync, inflateRawSync } from 'node:zlib'
import { unzipSync, zipSync } from 'fflate'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'
import { createZip } from '../../packages/io-xlsx/src/writer/zip.ts'

interface Sample {
	readonly name: string
	readonly customMs: number
	readonly nodeInflateMs: number
	readonly bunInflateMs: number | null
	readonly fflateMs: number
	readonly speedup: number
	readonly createZipMs: number
	readonly tableCrcMs: number
	readonly bunCrcMs: number | null
	readonly nodeDeflateMs: number
	readonly bunDeflateMs: number | null
	readonly fflateZipMs: number
	readonly deflateMatrix: readonly DeflateMatrixSample[]
	readonly partPhases: readonly PartPhaseSample[]
}

interface DeflateMatrixSample {
	readonly name: string
	readonly ms: number
	readonly bytes: number
}

interface PartPhaseSample {
	readonly fixture: string
	readonly path: string
	readonly compressedBytes: number
	readonly uncompressedBytes: number
	readonly nodeInflateMs: number | null
	readonly bunInflateMs: number | null
	readonly decodeMs: number
	readonly stringScanMs: number
	readonly byteScanMs: number
	readonly textTokens: number
	readonly byteTokens: number
}

interface DeflateCandidate {
	readonly name: string
	readonly level: number
	readonly memLevel?: number
}

interface FixtureSpec {
	readonly name: string
	readonly path: string
}

type BenchmarkPhase = 'all' | 'parts'

interface PartInput {
	readonly path: string
	readonly compressionMethod: number
	readonly compressed: Uint8Array
	readonly raw: Uint8Array
}

const FIXTURE_FILES = [
	'../xlsx/stress/dense-100k.xlsx',
	'../xlsx/stress/many-strings.xlsx',
	'../xlsx/poi/ConditionalFormattingSamples.xlsx',
] as const
const MAX_PHASE_PARTS_PER_FIXTURE = 6
const MIN_PHASE_PART_BYTES = 64 * 1024

const BUN_DEFLATE_CANDIDATES: readonly DeflateCandidate[] = [
	{ name: 'level2-production', level: 2 },
	{ name: 'level1-default', level: 1 },
	{ name: 'level1-mem9', level: 1, memLevel: 9 },
	{ name: 'level2-mem9', level: 2, memLevel: 9 },
	{ name: 'level3-mem9', level: 3, memLevel: 9 },
	{ name: 'level0-store-ish', level: 0, memLevel: 9 },
]

function main(): void {
	const baseDir = fileURLToPath(new URL('.', import.meta.url))
	const { repeat, fixtures, phase } = parseArgs(process.argv.slice(2), baseDir)
	const samples: Sample[] = []

	for (const fixture of fixtures) {
		const bytes = new Uint8Array(readFileSync(fixture.path))
		if (phase === 'all') warmup(bytes)
		const archive = extractZip(bytes)
		const compressedEntries = [...archive.entries()].map((entry) => ({
			entry,
			compressed: bytes.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize),
		}))
		const rawParts = new Map<string, Uint8Array>()
		for (const entry of archive.entries()) {
			const raw = archive.readBytes(entry.path)
			if (raw) rawParts.set(entry.path, raw)
		}
		const partInputs = selectPhaseParts(
			compressedEntries.map(({ entry, compressed }) => ({
				path: entry.path,
				compressionMethod: entry.compressionMethod,
				compressed,
				raw: rawParts.get(entry.path),
			})),
		)
		warmupPartPhases(partInputs)
		const partPhases = partInputs.map((part) => measurePartPhase(fixture.name, part, repeat))
		if (phase === 'parts') {
			samples.push({
				name: fixture.name,
				customMs: 0,
				nodeInflateMs: 0,
				bunInflateMs: null,
				fflateMs: 0,
				speedup: 0,
				createZipMs: 0,
				tableCrcMs: 0,
				bunCrcMs: null,
				nodeDeflateMs: 0,
				bunDeflateMs: null,
				fflateZipMs: 0,
				deflateMatrix: [],
				partPhases,
			})
			continue
		}
		const customMs = time(() => {
			const archive = extractZip(bytes)
			for (const entry of archive.entries()) {
				archive.readBytes(entry.path)
			}
		}, repeat)
		const nodeInflateMs = time(() => {
			for (const { entry, compressed } of compressedEntries) {
				if (entry.compressionMethod === 8) inflateRawSync(compressed)
			}
		}, repeat)
		const bunInflateMs = bunInflateRawSync
			? time(() => {
					for (const { entry, compressed } of compressedEntries) {
						if (entry.compressionMethod === 8) {
							bunInflateRawSync(compressed, { windowBits: -15 })
						}
					}
				}, repeat)
			: null
		const fflateMs = time(() => {
			unzipSync(bytes)
		}, repeat)
		const createZipMs = time(() => {
			createZip(rawParts)
		}, repeat)
		const tableCrcMs = time(() => {
			let checksum = 0
			for (const raw of rawParts.values()) {
				checksum ^= tableCrc32(raw)
			}
			consume(checksum)
		}, repeat)
		const bunCrcMs = bunCrc32
			? time(() => {
					let checksum = 0
					for (const raw of rawParts.values()) {
						checksum ^= bunCrc32(raw) >>> 0
					}
					consume(checksum)
				}, repeat)
			: null
		const nodeDeflateMs = time(() => {
			for (const [path, raw] of rawParts) {
				if (raw.byteLength >= 1024) {
					deflateRawSync(raw, { level: deflateLevelForPart(path, raw.byteLength) })
				}
			}
		}, repeat)
		const bunDeflateMs = bunDeflateRawSync
			? time(() => {
					for (const [path, raw] of rawParts) {
						if (raw.byteLength >= 1024) {
							bunDeflateRawSync(raw, {
								level: deflateLevelForPart(path, raw.byteLength),
								windowBits: -15,
							})
						}
					}
				}, repeat)
			: null
		const fflateZipMs = time(() => {
			zipSync(Object.fromEntries(rawParts))
		}, repeat)
		const worksheetParts = selectWorksheetParts(rawParts)
		const deflateMatrix = bunDeflateRawSync
			? BUN_DEFLATE_CANDIDATES.map((candidate) =>
					measureBunDeflateCandidate(worksheetParts, candidate, repeat),
				)
			: []
		samples.push({
			name: fixture.name,
			customMs,
			nodeInflateMs,
			bunInflateMs,
			fflateMs,
			speedup: customMs / fflateMs,
			createZipMs,
			tableCrcMs,
			bunCrcMs,
			nodeDeflateMs,
			bunDeflateMs,
			fflateZipMs,
			deflateMatrix,
			partPhases,
		})
	}

	if (phase === 'all') {
		console.log('ZIP Extraction A/B')
		console.log('='.repeat(88))
		console.log(
			[
				'Fixture'.padEnd(42),
				'Current(ms)'.padStart(12),
				'node inflate'.padStart(12),
				'Bun inflate'.padStart(12),
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
					sample.nodeInflateMs.toFixed(2).padStart(12),
					(sample.bunInflateMs === null ? 'n/a' : sample.bunInflateMs.toFixed(2)).padStart(12),
					sample.fflateMs.toFixed(2).padStart(12),
					`${sample.speedup.toFixed(2)}x`.padStart(10),
				].join(' '),
			)
		}
		console.log()
		console.log('ZIP CRC A/B')
		console.log('='.repeat(88))
		console.log(
			[
				'Fixture'.padEnd(42),
				'table crc'.padStart(12),
				'Bun crc'.padStart(12),
				'Bun/table'.padStart(12),
			].join(' '),
		)
		console.log('-'.repeat(88))
		for (const sample of samples) {
			console.log(
				[
					sample.name.padEnd(42),
					sample.tableCrcMs.toFixed(2).padStart(12),
					(sample.bunCrcMs === null ? 'n/a' : sample.bunCrcMs.toFixed(2)).padStart(12),
					(sample.bunCrcMs === null
						? 'n/a'
						: `${(sample.bunCrcMs / sample.tableCrcMs).toFixed(2)}x`
					).padStart(12),
				].join(' '),
			)
		}
		console.log()
		console.log('ZIP Creation A/B')
		console.log('='.repeat(88))
		console.log(
			[
				'Fixture'.padEnd(42),
				'createZip(ms)'.padStart(12),
				'node deflate'.padStart(12),
				'Bun deflate'.padStart(12),
				'fflate zip'.padStart(12),
			].join(' '),
		)
		console.log('-'.repeat(88))
		for (const sample of samples) {
			console.log(
				[
					sample.name.padEnd(42),
					sample.createZipMs.toFixed(2).padStart(12),
					sample.nodeDeflateMs.toFixed(2).padStart(12),
					(sample.bunDeflateMs === null ? 'n/a' : sample.bunDeflateMs.toFixed(2)).padStart(12),
					sample.fflateZipMs.toFixed(2).padStart(12),
				].join(' '),
			)
		}
		console.log()
		console.log('Worksheet Deflate Option A/B (Bun raw deflate)')
		console.log('='.repeat(88))
		for (const sample of samples) {
			if (sample.deflateMatrix.length === 0) continue
			const bestMs = Math.min(...sample.deflateMatrix.map((candidate) => candidate.ms))
			const smallestBytes = Math.min(...sample.deflateMatrix.map((candidate) => candidate.bytes))
			console.log(sample.name)
			console.log(
				[
					'Candidate'.padEnd(24),
					'ms'.padStart(10),
					'bytes'.padStart(12),
					'vs fastest'.padStart(12),
					'vs smallest'.padStart(12),
				].join(' '),
			)
			for (const candidate of sample.deflateMatrix) {
				console.log(
					[
						candidate.name.padEnd(24),
						candidate.ms.toFixed(2).padStart(10),
						String(candidate.bytes).padStart(12),
						`${(candidate.ms / bestMs).toFixed(2)}x`.padStart(12),
						`${(candidate.bytes / smallestBytes).toFixed(2)}x`.padStart(12),
					].join(' '),
				)
			}
			console.log()
		}
	}
	const phaseRows = samples.flatMap((sample) => sample.partPhases)
	if (phaseRows.length > 0) {
		console.log('Large XML Part Phase A/B')
		console.log('='.repeat(132))
		console.log(
			[
				'Fixture'.padEnd(22),
				'Part'.padEnd(42),
				'raw MB'.padStart(8),
				'zip MB'.padStart(8),
				'node infl'.padStart(10),
				'Bun infl'.padStart(10),
				'decode'.padStart(9),
				'str scan'.padStart(9),
				'byte scan'.padStart(10),
				'byte/str'.padStart(9),
			].join(' '),
		)
		console.log('-'.repeat(132))
		for (const row of phaseRows) {
			console.log(
				[
					truncate(row.fixture, 22).padEnd(22),
					truncate(row.path, 42).padEnd(42),
					(row.uncompressedBytes / 1024 / 1024).toFixed(1).padStart(8),
					(row.compressedBytes / 1024 / 1024).toFixed(1).padStart(8),
					formatNullableMs(row.nodeInflateMs).padStart(10),
					formatNullableMs(row.bunInflateMs).padStart(10),
					row.decodeMs.toFixed(2).padStart(9),
					row.stringScanMs.toFixed(2).padStart(9),
					row.byteScanMs.toFixed(2).padStart(10),
					`${(row.byteScanMs / row.stringScanMs).toFixed(2)}x`.padStart(9),
				].join(' '),
			)
		}
		console.log()
	}
}

function parseArgs(
	argv: readonly string[],
	baseDir: string,
): {
	readonly repeat: number
	readonly fixtures: readonly FixtureSpec[]
	readonly phase: BenchmarkPhase
} {
	let repeat = 8
	let phase: BenchmarkPhase = 'all'
	const fixturePaths: string[] = []
	for (let i = 0; i < argv.length; i++) {
		const arg = argv[i]
		if (arg === '--repeat' || arg === '-r') {
			const value = argv[++i]
			if (!value) throw new Error(`Missing value for ${arg}`)
			repeat = Math.max(1, Number.parseInt(value, 10) || repeat)
			continue
		}
		if (arg === '--fixture') {
			const value = argv[++i]
			if (!value) throw new Error(`Missing value for ${arg}`)
			fixturePaths.push(value)
			continue
		}
		if (arg === '--phase') {
			const value = argv[++i]
			if (value !== 'all' && value !== 'parts') {
				throw new Error('Unsupported --phase. Expected "all" or "parts"')
			}
			phase = value
			continue
		}
		if (arg && /^\d+$/.test(arg)) {
			repeat = Math.max(1, Number.parseInt(arg, 10) || repeat)
			continue
		}
		if (arg) fixturePaths.push(arg)
	}
	const paths = fixturePaths.length > 0 ? fixturePaths : [...FIXTURE_FILES]
	return {
		repeat,
		phase,
		fixtures: paths.map((path) => {
			const resolvedPath = resolveFixturePath(path, baseDir)
			return { name: displayFixtureName(path), path: resolvedPath }
		}),
	}
}

function resolveFixturePath(path: string, baseDir: string): string {
	if (isAbsolute(path)) return path
	if (path.startsWith('../')) return join(baseDir, path)
	return resolve(process.cwd(), path)
}

function displayFixtureName(path: string): string {
	return path.replace('../xlsx/', '').replace(/^.*\/research\/excel-corpus\//, 'research/')
}

function selectPhaseParts(
	entries: readonly (Omit<PartInput, 'raw'> & { readonly raw?: Uint8Array })[],
): readonly PartInput[] {
	const xmlEntries = entries
		.filter(
			(entry): entry is PartInput =>
				entry.raw !== undefined &&
				entry.path.endsWith('.xml') &&
				entry.raw.byteLength >= MIN_PHASE_PART_BYTES,
		)
		.sort((a, b) => b.raw.byteLength - a.raw.byteLength)
	if (xmlEntries.length > 0) return xmlEntries.slice(0, MAX_PHASE_PARTS_PER_FIXTURE)
	return entries
		.filter(
			(entry): entry is PartInput =>
				entry.raw !== undefined && entry.raw.byteLength >= MIN_PHASE_PART_BYTES,
		)
		.sort((a, b) => b.raw.byteLength - a.raw.byteLength)
		.slice(0, MAX_PHASE_PARTS_PER_FIXTURE)
}

function warmupPartPhases(parts: readonly PartInput[]): void {
	for (const part of parts) {
		if (part.compressionMethod === 8 && bunInflateRawSync) {
			consumeBytes(bunInflateRawSync(part.compressed, { windowBits: -15 }))
		}
		const text = new TextDecoder('utf-8').decode(part.raw)
		consume(scanXmlTextTokens(text))
		consume(scanXmlByteTokens(part.raw))
	}
}

function measurePartPhase(fixture: string, part: PartInput, repeat: number): PartPhaseSample {
	const decoder = new TextDecoder('utf-8')
	const text = decoder.decode(part.raw)
	const nodeInflateMs =
		part.compressionMethod === 8
			? time(() => {
					consumeBytes(inflateRawSync(part.compressed))
				}, repeat)
			: null
	const bunInflateMs =
		part.compressionMethod === 8 && bunInflateRawSync
			? time(() => {
					consumeBytes(bunInflateRawSync(part.compressed, { windowBits: -15 }))
				}, repeat)
			: null
	const decodeMs = time(() => {
		const decoded = decoder.decode(part.raw)
		consume(decoded.length)
	}, repeat)
	let textTokens = 0
	const stringScanMs = time(() => {
		textTokens = scanXmlTextTokens(text)
		consume(textTokens)
	}, repeat)
	let byteTokens = 0
	const byteScanMs = time(() => {
		byteTokens = scanXmlByteTokens(part.raw)
		consume(byteTokens)
	}, repeat)
	return {
		fixture,
		path: part.path,
		compressedBytes: part.compressed.byteLength,
		uncompressedBytes: part.raw.byteLength,
		nodeInflateMs,
		bunInflateMs,
		decodeMs,
		stringScanMs,
		byteScanMs,
		textTokens,
		byteTokens,
	}
}

function scanXmlTextTokens(text: string): number {
	let count = 0
	for (let i = 0; i < text.length - 1; i++) {
		if (text.charCodeAt(i) !== 60) continue
		const next = text.charCodeAt(i + 1)
		if (next === 99 || next === 118 || next === 102) {
			count++
			continue
		}
		if (next === 114 && text.charCodeAt(i + 2) === 111 && text.charCodeAt(i + 3) === 119) {
			count++
		}
	}
	return count
}

function scanXmlByteTokens(bytes: Uint8Array): number {
	let count = 0
	for (let i = 0; i < bytes.byteLength - 1; i++) {
		if (bytes[i] !== 60) continue
		const next = bytes[i + 1]
		if (next === 99 || next === 118 || next === 102) {
			count++
			continue
		}
		if (next === 114 && bytes[i + 2] === 111 && bytes[i + 3] === 119) {
			count++
		}
	}
	return count
}

function consumeBytes(bytes: Uint8Array): void {
	consume(bytes.byteLength ^ (bytes[0] ?? 0))
}

function formatNullableMs(value: number | null): string {
	return value === null ? 'n/a' : value.toFixed(2)
}

function truncate(value: string, maxLength: number): string {
	if (value.length <= maxLength) return value
	return `...${value.slice(Math.max(0, value.length - maxLength + 3))}`
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

type BunInflateSync = (data: Uint8Array, options: { readonly windowBits: number }) => Uint8Array
type BunDeflateSync = (
	data: Uint8Array,
	options: { readonly level: number; readonly windowBits: number; readonly memLevel?: number },
) => Uint8Array
type BunCrc32 = (data: Uint8Array) => number

const bunInflateRawSync = (() => {
	const maybeBun = (globalThis as { readonly Bun?: { readonly inflateSync?: BunInflateSync } }).Bun
	return typeof maybeBun?.inflateSync === 'function' ? maybeBun.inflateSync.bind(maybeBun) : null
})()

const bunDeflateRawSync = (() => {
	const maybeBun = (globalThis as { readonly Bun?: { readonly deflateSync?: BunDeflateSync } }).Bun
	return typeof maybeBun?.deflateSync === 'function' ? maybeBun.deflateSync.bind(maybeBun) : null
})()

const bunCrc32 = (() => {
	const maybeBun = (
		globalThis as { readonly Bun?: { readonly hash?: { readonly crc32?: BunCrc32 } } }
	).Bun
	return typeof maybeBun?.hash?.crc32 === 'function'
		? maybeBun.hash.crc32.bind(maybeBun.hash)
		: null
})()

const CRC_TABLES = /* @__PURE__ */ (() => {
	const t = new Uint32Array(256)
	for (let i = 0; i < 256; i++) {
		let c = i
		for (let j = 0; j < 8; j++) {
			c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
		}
		t[i] = c
	}
	const tables = [t]
	for (let tableIndex = 1; tableIndex < 8; tableIndex++) {
		const previous = tables[tableIndex - 1] as Uint32Array
		const next = new Uint32Array(256)
		for (let i = 0; i < 256; i++) {
			const c = previous[i] as number
			next[i] = ((t[c & 0xff] as number) ^ (c >>> 8)) >>> 0
		}
		tables.push(next)
	}
	return tables as readonly Uint32Array[]
})()

let _consumedChecksum = 0

function consume(value: number): void {
	_consumedChecksum ^= value
}

function tableCrc32(data: Uint8Array): number {
	return (updateTableCrc32(0xffffffff, data) ^ 0xffffffff) >>> 0
}

function updateTableCrc32(crc: number, data: Uint8Array): number {
	const t0 = CRC_TABLES[0] as Uint32Array
	const t1 = CRC_TABLES[1] as Uint32Array
	const t2 = CRC_TABLES[2] as Uint32Array
	const t3 = CRC_TABLES[3] as Uint32Array
	const t4 = CRC_TABLES[4] as Uint32Array
	const t5 = CRC_TABLES[5] as Uint32Array
	const t6 = CRC_TABLES[6] as Uint32Array
	const t7 = CRC_TABLES[7] as Uint32Array
	let i = 0
	const length = data.byteLength
	const blockLength = length - (length % 8)
	while (i < blockLength) {
		const word =
			((data[i] as number) |
				((data[i + 1] as number) << 8) |
				((data[i + 2] as number) << 16) |
				((data[i + 3] as number) << 24)) ^
			crc
		const next =
			(data[i + 4] as number) |
			((data[i + 5] as number) << 8) |
			((data[i + 6] as number) << 16) |
			((data[i + 7] as number) << 24)
		crc =
			((t7[word & 0xff] as number) ^
				(t6[(word >>> 8) & 0xff] as number) ^
				(t5[(word >>> 16) & 0xff] as number) ^
				(t4[(word >>> 24) & 0xff] as number) ^
				(t3[next & 0xff] as number) ^
				(t2[(next >>> 8) & 0xff] as number) ^
				(t1[(next >>> 16) & 0xff] as number) ^
				(t0[(next >>> 24) & 0xff] as number)) >>>
			0
		i += 8
	}
	while (i < length) {
		crc = ((t0[(crc ^ (data[i] as number)) & 0xff] as number) ^ (crc >>> 8)) >>> 0
		i++
	}
	return crc
}

function deflateLevelForPart(path: string, uncompressedSize: number): number {
	if (/xl\/worksheets\/sheet\d+\.xml$/i.test(path)) return 2
	if (uncompressedSize > 512_000) return 1
	if (uncompressedSize < 32_000) return 6
	return 2
}

function selectWorksheetParts(rawParts: ReadonlyMap<string, Uint8Array>): readonly Uint8Array[] {
	const worksheetParts = [...rawParts]
		.filter(([path, raw]) => /xl\/worksheets\/sheet\d+\.xml$/i.test(path) && raw.byteLength >= 1024)
		.map(([, raw]) => raw)
	if (worksheetParts.length > 0) return worksheetParts
	return [...rawParts.values()].filter((raw) => raw.byteLength >= 1024)
}

function measureBunDeflateCandidate(
	parts: readonly Uint8Array[],
	candidate: DeflateCandidate,
	repeat: number,
): DeflateMatrixSample {
	let bytes = 0
	const ms = time(() => {
		let totalBytes = 0
		for (const raw of parts) {
			const compressed = bunDeflateRawSync?.(raw, {
				level: candidate.level,
				windowBits: -15,
				...(candidate.memLevel === undefined ? {} : { memLevel: candidate.memLevel }),
			})
			totalBytes += compressed?.byteLength ?? 0
		}
		bytes = totalBytes
	}, repeat)
	return { name: candidate.name, ms, bytes }
}

main()
