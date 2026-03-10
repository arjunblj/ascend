import { inflateRawSync } from 'node:zlib'

const EOCD_SIGNATURE = 0x06054b50
const CENTRAL_DIR_SIGNATURE = 0x02014b50
const LOCAL_FILE_SIGNATURE = 0x04034b50
const MAX_CACHED_PART_BYTES = 128 * 1024

export interface ZipEntry {
	readonly path: string
	readonly compressionMethod: number
	readonly compressedSize: number
	readonly uncompressedSize: number
	readonly localHeaderOffset: number
	readonly dataOffset: number
}

export class ZipArchive {
	private readonly bytes: Uint8Array
	private readonly entriesByPath: Map<string, ZipEntry>
	private readonly decoder = new TextDecoder('utf-8')
	private readonly bytesCache = new Map<string, Uint8Array>()
	private readonly textCache = new Map<string, string>()

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
		else if (entry.compressionMethod === 8) bytes = new Uint8Array(inflateRawSync(compressed))
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
		}
		return text
	}
}

export function extractZip(bytes: Uint8Array): ZipArchive {
	return new ZipArchive(bytes)
}

function parseEntries(bytes: Uint8Array, decoder: TextDecoder): Map<string, ZipEntry> {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const eocdOffset = findEocdOffset(view)
	const entryCount = view.getUint16(eocdOffset + 10, true)
	const centralDirOffset = view.getUint32(eocdOffset + 16, true)
	const centralDirSize = view.getUint32(eocdOffset + 12, true)

	if (centralDirOffset === 0xffffffff || centralDirSize === 0xffffffff) {
		throw new Error('ZIP64 archives are not supported yet')
	}

	const entries = new Map<string, ZipEntry>()
	let offset = centralDirOffset
	for (let i = 0; i < entryCount; i++) {
		if (view.getUint32(offset, true) !== CENTRAL_DIR_SIGNATURE) {
			throw new Error('Invalid central directory entry')
		}

		const compressionMethod = view.getUint16(offset + 10, true)
		const compressedSize = view.getUint32(offset + 20, true)
		const uncompressedSize = view.getUint32(offset + 24, true)
		const fileNameLength = view.getUint16(offset + 28, true)
		const extraFieldLength = view.getUint16(offset + 30, true)
		const fileCommentLength = view.getUint16(offset + 32, true)
		const localHeaderOffset = view.getUint32(offset + 42, true)
		if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff) {
			throw new Error('ZIP64 entries are not supported yet')
		}

		const fileNameStart = offset + 46
		const fileNameEnd = fileNameStart + fileNameLength
		const path = decoder.decode(bytes.subarray(fileNameStart, fileNameEnd))
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
