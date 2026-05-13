import type { Buffer } from 'node:buffer'
import { Readable } from 'node:stream'
import { promisify } from 'node:util'
import { createInflateRaw, inflateRaw, inflateRawSync } from 'node:zlib'
import { Inflate } from 'fflate'

const EOCD_SIGNATURE = 0x06054b50
const ZIP64_EOCD_SIGNATURE = 0x06064b50
const ZIP64_EOCD_LOCATOR_SIGNATURE = 0x07064b50
const CENTRAL_DIR_SIGNATURE = 0x02014b50
const LOCAL_FILE_SIGNATURE = 0x04034b50
const ZIP64_EXTRA_FIELD_ID = 0x0001
const MAX_CACHED_PART_BYTES = 128 * 1024
const STREAM_CHUNK_BYTES = 1024 * 1024
const FULL_INFLATE_TEXT_CHUNK_LIMIT_BYTES = 64 * 1024 * 1024
const SYNC_INFLATE_ASYNC_LIMIT_BYTES = 8 * 1024 * 1024
const inflateRawAsync = promisify(inflateRaw)
type BunInflateSync = (data: Uint8Array, options: { readonly windowBits: number }) => Uint8Array
const bunInflateRawSync = getBunInflateRawSync()

export interface ZipEntry {
	readonly path: string
	readonly compressionMethod: number
	readonly compressedSize: number
	readonly uncompressedSize: number
	readonly localHeaderOffset: number
	readonly dataOffset: number
}

interface ReadChunksOptions {
	readonly preferStreaming?: boolean
}

export class ZipArchive {
	private readonly bytes: Uint8Array
	private readonly entriesByPath: Map<string, ZipEntry>
	private readonly decoder = new TextDecoder('utf-8')
	private readonly bytesCache = new Map<string, Uint8Array>()
	private readonly textCache = new Map<string, string>()
	private caseInsensitivePathByLower: Map<string, string> | undefined

	constructor(bytes: Uint8Array) {
		this.bytes = bytes
		this.entriesByPath = parseEntries(bytes, this.decoder)
	}

	get(path: string): ZipEntry | undefined {
		return this.entriesByPath.get(path)
	}

	has(path: string): boolean {
		return this.entriesByPath.has(path)
	}

	entries(): IterableIterator<ZipEntry> {
		return this.entriesByPath.values()
	}

	resolvePathCaseInsensitive(path: string): string | undefined {
		this.caseInsensitivePathByLower ??= buildCaseInsensitivePathMap(this.entriesByPath)
		return this.caseInsensitivePathByLower.get(path.toLowerCase())
	}

	readBytes(path: string): Uint8Array | undefined {
		const cached = this.bytesCache.get(path)
		if (cached) return cached
		const entry = this.entriesByPath.get(path)
		if (!entry) return undefined
		const compressed = this.bytes.subarray(
			entry.dataOffset,
			entry.dataOffset + entry.compressedSize,
		)
		let bytes: Uint8Array
		if (entry.compressionMethod === 0) bytes = compressed
		else if (entry.compressionMethod === 8) bytes = inflateRawBytesSync(compressed)
		else
			throw new Error(`Unsupported ZIP compression method ${entry.compressionMethod} for ${path}`)
		if (bytes.byteLength <= MAX_CACHED_PART_BYTES) {
			this.bytesCache.set(path, bytes)
		}
		return bytes
	}

	readText(path: string): string | undefined {
		const cached = this.textCache.get(path)
		if (cached !== undefined) return cached
		const bytes = this.readBytes(path)
		if (!bytes) return undefined
		const text = this.decoder.decode(bytes)
		if (bytes.byteLength <= MAX_CACHED_PART_BYTES) {
			this.textCache.set(path, text)
			this.bytesCache.delete(path)
		}
		return text
	}

	*readTextChunks(path: string, chunkSize = STREAM_CHUNK_BYTES): IterableIterator<string> {
		const entry = this.entriesByPath.get(path)
		if (!entry) return
		const compressed = this.bytes.subarray(
			entry.dataOffset,
			entry.dataOffset + entry.compressedSize,
		)
		const decoder = new TextDecoder('utf-8')
		if (entry.compressionMethod === 0) {
			yield* decodeTextChunks(compressed, decoder, chunkSize)
			return
		}
		if (entry.compressionMethod !== 8) {
			throw new Error(`Unsupported ZIP compression method ${entry.compressionMethod} for ${path}`)
		}
		if (bunInflateRawSync && entry.uncompressedSize <= FULL_INFLATE_TEXT_CHUNK_LIMIT_BYTES) {
			yield* decodeTextChunks(inflateRawBytesSync(compressed), decoder, chunkSize)
			return
		}
		const pending: string[] = []
		const inflater = new Inflate((chunk, final) => {
			const text = decoder.decode(chunk, { stream: !final })
			if (text) pending.push(text)
			if (final) {
				const tail = decoder.decode()
				if (tail) pending.push(tail)
			}
		})
		for (let offset = 0; offset < compressed.byteLength; offset += chunkSize) {
			inflater.push(
				compressed.subarray(offset, offset + chunkSize),
				offset + chunkSize >= compressed.byteLength,
			)
			while (pending.length > 0) {
				const text = pending.shift()
				if (text) yield text
			}
		}
	}

	*readByteChunks(
		path: string,
		chunkSize = STREAM_CHUNK_BYTES,
		options: ReadChunksOptions = {},
	): IterableIterator<Uint8Array> {
		const entry = this.entriesByPath.get(path)
		if (!entry) return
		const compressed = this.bytes.subarray(
			entry.dataOffset,
			entry.dataOffset + entry.compressedSize,
		)
		if (entry.compressionMethod === 0) {
			for (let offset = 0; offset < compressed.byteLength; offset += chunkSize) {
				yield compressed.subarray(offset, offset + chunkSize)
			}
			return
		}
		if (entry.compressionMethod !== 8) {
			throw new Error(`Unsupported ZIP compression method ${entry.compressionMethod} for ${path}`)
		}
		if (
			!options.preferStreaming &&
			bunInflateRawSync &&
			entry.uncompressedSize <= FULL_INFLATE_TEXT_CHUNK_LIMIT_BYTES
		) {
			const bytes = inflateRawBytesSync(compressed)
			for (let offset = 0; offset < bytes.byteLength; offset += chunkSize) {
				yield bytes.subarray(offset, offset + chunkSize)
			}
			return
		}
		const pending: Uint8Array[] = []
		const inflater = new Inflate((chunk) => {
			if (chunk.byteLength > 0) pending.push(chunk)
		})
		for (let offset = 0; offset < compressed.byteLength; offset += chunkSize) {
			inflater.push(
				compressed.subarray(offset, offset + chunkSize),
				offset + chunkSize >= compressed.byteLength,
			)
			while (pending.length > 0) {
				const bytes = pending.shift()
				if (bytes) yield bytes
			}
		}
	}

	async *readTextChunksAsync(path: string, chunkSize = STREAM_CHUNK_BYTES): AsyncGenerator<string> {
		const entry = this.entriesByPath.get(path)
		if (!entry) return
		const compressed = this.bytes.subarray(
			entry.dataOffset,
			entry.dataOffset + entry.compressedSize,
		)
		const decoder = new TextDecoder('utf-8')
		if (entry.compressionMethod === 0) {
			for (let offset = 0; offset < compressed.byteLength; offset += chunkSize) {
				const text = decoder.decode(compressed.subarray(offset, offset + chunkSize), {
					stream: offset + chunkSize < compressed.byteLength,
				})
				if (text) yield text
			}
			const tail = decoder.decode()
			if (tail) yield tail
			return
		}
		if (entry.compressionMethod !== 8) {
			throw new Error(`Unsupported ZIP compression method ${entry.compressionMethod} for ${path}`)
		}
		const source = Readable.from(
			(function* () {
				for (let offset = 0; offset < compressed.byteLength; offset += chunkSize) {
					yield compressed.subarray(offset, offset + chunkSize)
				}
			})(),
		)
		for await (const chunk of source.pipe(createInflateRaw())) {
			const text = decoder.decode(chunk as Uint8Array, { stream: true })
			if (text) yield text
		}
		const tail = decoder.decode()
		if (tail) yield tail
	}

	async *readByteChunksAsync(
		path: string,
		chunkSize = STREAM_CHUNK_BYTES,
	): AsyncGenerator<Uint8Array> {
		const entry = this.entriesByPath.get(path)
		if (!entry) return
		const compressed = this.bytes.subarray(
			entry.dataOffset,
			entry.dataOffset + entry.compressedSize,
		)
		if (entry.compressionMethod === 0) {
			for (let offset = 0; offset < compressed.byteLength; offset += chunkSize) {
				yield compressed.subarray(offset, offset + chunkSize)
			}
			return
		}
		if (entry.compressionMethod !== 8) {
			throw new Error(`Unsupported ZIP compression method ${entry.compressionMethod} for ${path}`)
		}
		const source = Readable.from(
			(function* () {
				for (let offset = 0; offset < compressed.byteLength; offset += chunkSize) {
					yield compressed.subarray(offset, offset + chunkSize)
				}
			})(),
		)
		for await (const chunk of source.pipe(createInflateRaw())) {
			yield chunk as Uint8Array
		}
	}

	async readBytesAsync(path: string): Promise<Uint8Array | undefined> {
		const cached = this.bytesCache.get(path)
		if (cached) return cached
		const entry = this.entriesByPath.get(path)
		if (!entry) return undefined
		const compressed = this.bytes.subarray(
			entry.dataOffset,
			entry.dataOffset + entry.compressedSize,
		)
		let bytes: Uint8Array
		if (entry.compressionMethod === 0) bytes = compressed
		else if (entry.compressionMethod === 8) {
			bytes = asUint8Array(
				bunInflateRawSync && entry.uncompressedSize <= SYNC_INFLATE_ASYNC_LIMIT_BYTES
					? inflateRawBytesSync(compressed)
					: await inflateRawAsync(compressed),
			)
		} else
			throw new Error(`Unsupported ZIP compression method ${entry.compressionMethod} for ${path}`)
		if (bytes.byteLength <= MAX_CACHED_PART_BYTES) {
			this.bytesCache.set(path, bytes)
		}
		return bytes
	}

	async readTextAsync(path: string): Promise<string | undefined> {
		const cached = this.textCache.get(path)
		if (cached !== undefined) return cached
		const bytes = await this.readBytesAsync(path)
		if (!bytes) return undefined
		const text = this.decoder.decode(bytes)
		if (bytes.byteLength <= MAX_CACHED_PART_BYTES) {
			this.textCache.set(path, text)
		}
		return text
	}

	async readTextsAsync(paths: readonly string[]): Promise<Map<string, string>> {
		const entries = await Promise.all(
			paths.map(async (path) => {
				const text = await this.readTextAsync(path)
				return text === undefined ? null : ([path, text] as const)
			}),
		)
		const results = new Map<string, string>()
		for (const entry of entries) {
			if (!entry) continue
			results.set(entry[0], entry[1])
		}
		return results
	}
}

export function extractZip(bytes: Uint8Array): ZipArchive {
	return new ZipArchive(bytes)
}

function buildCaseInsensitivePathMap(entries: ReadonlyMap<string, ZipEntry>): Map<string, string> {
	const byLower = new Map<string, string>()
	for (const path of entries.keys()) {
		const lower = path.toLowerCase()
		if (!byLower.has(lower)) byLower.set(lower, path)
	}
	return byLower
}

function inflateRawBytesSync(compressed: Uint8Array): Uint8Array {
	return asUint8Array(
		bunInflateRawSync?.(compressed, { windowBits: -15 }) ?? inflateRawSync(compressed),
	)
}

function asUint8Array(bytes: Uint8Array | Buffer): Uint8Array {
	return bytes.constructor === Uint8Array
		? bytes
		: new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function* decodeTextChunks(
	bytes: Uint8Array,
	decoder: TextDecoder,
	chunkSize: number,
): IterableIterator<string> {
	for (let offset = 0; offset < bytes.byteLength; offset += chunkSize) {
		const text = decoder.decode(bytes.subarray(offset, offset + chunkSize), {
			stream: offset + chunkSize < bytes.byteLength,
		})
		if (text) yield text
	}
	const tail = decoder.decode()
	if (tail) yield tail
}

function getBunInflateRawSync(): BunInflateSync | undefined {
	const maybeBun = (globalThis as { readonly Bun?: { readonly inflateSync?: BunInflateSync } }).Bun
	if (typeof maybeBun?.inflateSync !== 'function') return undefined
	return maybeBun.inflateSync.bind(maybeBun)
}

function parseEntries(bytes: Uint8Array, decoder: TextDecoder): Map<string, ZipEntry> {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const eocdOffset = findEocdOffset(view)
	const directory = readCentralDirectoryInfo(view, eocdOffset)

	const entries = new Map<string, ZipEntry>()
	let offset = directory.centralDirOffset
	for (let i = 0; i < directory.entryCount; i++) {
		if (view.getUint32(offset, true) !== CENTRAL_DIR_SIGNATURE) {
			throw new Error('Invalid central directory entry')
		}

		const compressionMethod = view.getUint16(offset + 10, true)
		let compressedSize = view.getUint32(offset + 20, true)
		let uncompressedSize = view.getUint32(offset + 24, true)
		const fileNameLength = view.getUint16(offset + 28, true)
		const extraFieldLength = view.getUint16(offset + 30, true)
		const fileCommentLength = view.getUint16(offset + 32, true)
		let localHeaderOffset = view.getUint32(offset + 42, true)

		const fileNameStart = offset + 46
		const fileNameEnd = fileNameStart + fileNameLength
		const extraFieldStart = fileNameEnd
		const extraFieldEnd = extraFieldStart + extraFieldLength
		const path = decoder.decode(bytes.subarray(fileNameStart, fileNameEnd)).replaceAll('\\', '/')
		if (
			compressedSize === 0xffffffff ||
			uncompressedSize === 0xffffffff ||
			localHeaderOffset === 0xffffffff
		) {
			const zip64 = parseZip64ExtraField(bytes.subarray(extraFieldStart, extraFieldEnd), {
				uncompressedSize: uncompressedSize === 0xffffffff,
				compressedSize: compressedSize === 0xffffffff,
				localHeaderOffset: localHeaderOffset === 0xffffffff,
			})
			uncompressedSize = zip64.uncompressedSize ?? uncompressedSize
			compressedSize = zip64.compressedSize ?? compressedSize
			localHeaderOffset = zip64.localHeaderOffset ?? localHeaderOffset
		}
		const dataOffset = getDataOffset(view, localHeaderOffset)
		entries.set(path, {
			path,
			compressionMethod,
			compressedSize,
			uncompressedSize,
			localHeaderOffset,
			dataOffset,
		})

		offset = fileNameEnd + extraFieldLength + fileCommentLength
	}

	return entries
}

function readCentralDirectoryInfo(
	view: DataView,
	eocdOffset: number,
): {
	entryCount: number
	centralDirOffset: number
	centralDirSize: number
} {
	const entryCount = view.getUint16(eocdOffset + 10, true)
	const centralDirOffset = view.getUint32(eocdOffset + 16, true)
	const centralDirSize = view.getUint32(eocdOffset + 12, true)
	if (entryCount !== 0xffff && centralDirOffset !== 0xffffffff && centralDirSize !== 0xffffffff) {
		return { entryCount, centralDirOffset, centralDirSize }
	}
	return readZip64CentralDirectoryInfo(view, eocdOffset)
}

function readZip64CentralDirectoryInfo(
	view: DataView,
	eocdOffset: number,
): {
	entryCount: number
	centralDirOffset: number
	centralDirSize: number
} {
	const locatorOffset = eocdOffset - 20
	if (locatorOffset < 0 || view.getUint32(locatorOffset, true) !== ZIP64_EOCD_LOCATOR_SIGNATURE) {
		throw new Error('Missing ZIP64 end of central directory locator')
	}
	const zip64EocdOffset = readUint64(view, locatorOffset + 8)
	if (view.getUint32(zip64EocdOffset, true) !== ZIP64_EOCD_SIGNATURE) {
		throw new Error('Invalid ZIP64 end of central directory record')
	}
	return {
		entryCount: readUint64(view, zip64EocdOffset + 32),
		centralDirSize: readUint64(view, zip64EocdOffset + 40),
		centralDirOffset: readUint64(view, zip64EocdOffset + 48),
	}
}

function parseZip64ExtraField(
	extra: Uint8Array,
	required: {
		uncompressedSize: boolean
		compressedSize: boolean
		localHeaderOffset: boolean
	},
): {
	uncompressedSize?: number
	compressedSize?: number
	localHeaderOffset?: number
} {
	const view = new DataView(extra.buffer, extra.byteOffset, extra.byteLength)
	let offset = 0
	while (offset + 4 <= view.byteLength) {
		const headerId = view.getUint16(offset, true)
		const dataSize = view.getUint16(offset + 2, true)
		offset += 4
		if (offset + dataSize > view.byteLength) break
		if (headerId === ZIP64_EXTRA_FIELD_ID) {
			let fieldOffset = offset
			const result: {
				uncompressedSize?: number
				compressedSize?: number
				localHeaderOffset?: number
			} = {}
			if (required.uncompressedSize) {
				result.uncompressedSize = readUint64(view, fieldOffset)
				fieldOffset += 8
			}
			if (required.compressedSize) {
				result.compressedSize = readUint64(view, fieldOffset)
				fieldOffset += 8
			}
			if (required.localHeaderOffset) {
				result.localHeaderOffset = readUint64(view, fieldOffset)
			}
			return result
		}
		offset += dataSize
	}
	throw new Error('Missing ZIP64 extended information extra field')
}

function findEocdOffset(view: DataView): number {
	const minOffset = Math.max(0, view.byteLength - 65_557)
	for (let offset = view.byteLength - 22; offset >= minOffset; offset--) {
		if (view.getUint32(offset, true) === EOCD_SIGNATURE) return offset
	}
	throw new Error('Missing end of central directory record')
}

function getDataOffset(view: DataView, localHeaderOffset: number): number {
	if (view.getUint32(localHeaderOffset, true) !== LOCAL_FILE_SIGNATURE) {
		throw new Error('Invalid local file header')
	}
	const fileNameLength = view.getUint16(localHeaderOffset + 26, true)
	const extraFieldLength = view.getUint16(localHeaderOffset + 28, true)
	return localHeaderOffset + 30 + fileNameLength + extraFieldLength
}

function readUint64(view: DataView, offset: number): number {
	const value = view.getBigUint64(offset, true)
	if (value > BigInt(Number.MAX_SAFE_INTEGER)) {
		throw new Error('ZIP64 value exceeds JavaScript safe integer range')
	}
	return Number(value)
}
