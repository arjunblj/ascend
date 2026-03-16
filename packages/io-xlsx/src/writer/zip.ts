import { createDeflateRaw, deflateRawSync } from 'node:zlib'

const textEncoder = new TextEncoder()
const DEFLATE_OPTS = { level: 2 } as const

const LOCAL_FILE_SIGNATURE = 0x04034b50
const CENTRAL_DIR_SIGNATURE = 0x02014b50
const EOCD_SIGNATURE = 0x06054b50
const ZIP64_EOCD_SIGNATURE = 0x06064b50
const ZIP64_EOCD_LOCATOR_SIGNATURE = 0x07064b50
const ZIP64_EXTRA_FIELD_ID = 0x0001
const ZIP64_VERSION = 45
const ZIP_VERSION = 20
const UINT16_MAX = 0xffff
const UINT32_MAX = 0xffffffff

const CRC_TABLE = /* @__PURE__ */ (() => {
	const t = new Uint32Array(256)
	for (let i = 0; i < 256; i++) {
		let c = i
		for (let j = 0; j < 8; j++) {
			c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
		}
		t[i] = c
	}
	return t
})()

function crc32(data: Uint8Array): number {
	let crc = 0xffffffff
	for (let i = 0; i < data.byteLength; i++) {
		const idx = (crc ^ (data[i] as number)) & 0xff
		crc = ((CRC_TABLE[idx] as number) ^ (crc >>> 8)) >>> 0
	}
	return (crc ^ 0xffffffff) >>> 0
}

interface CompressedEntry {
	nameBytes: Uint8Array
	data: Uint8Array
	uncompressedSize: number
	compressedSize: number
	method: number
	crc: number
}

export function createZip(parts: ReadonlyMap<string, Uint8Array>): Uint8Array {
	const entries: CompressedEntry[] = []

	for (const [path, raw] of parts) {
		const nameBytes = textEncoder.encode(path)
		const crc = crc32(raw)
		const deflated = new Uint8Array(deflateRawSync(raw, DEFLATE_OPTS))
		const useStore = deflated.byteLength >= raw.byteLength
		const data = useStore ? raw : deflated
		entries.push({
			nameBytes,
			data,
			uncompressedSize: raw.byteLength,
			compressedSize: data.byteLength,
			method: useStore ? 0 : 8,
			crc,
		})
	}
	return buildZip(entries)
}

interface ZipEntry {
	nameBytes: Uint8Array
	data: Uint8Array
	uncompressedSize: number
	compressedSize: number
	method: number
	crc: number
}

export class StreamingZipBuilder {
	private readonly entries: ZipEntry[] = []
	private streamingNameBytes: Uint8Array | null = null
	private streamingCrc = 0xffffffff
	private streamingUncompressedSize = 0
	private streamingCompressedChunks: Uint8Array[] = []
	private deflateStream: ReturnType<typeof createDeflateRaw> | null = null

	addEntry(path: string, data: Uint8Array): void {
		if (this.deflateStream) {
			throw new Error('Cannot addEntry while streaming entry is open; call closeEntry first')
		}
		const nameBytes = textEncoder.encode(path)
		const crc = crc32(data)
		const deflated = new Uint8Array(deflateRawSync(data, DEFLATE_OPTS))
		const useStore = deflated.byteLength >= data.byteLength
		const entryData = useStore ? data : deflated
		this.entries.push({
			nameBytes,
			data: entryData,
			uncompressedSize: data.byteLength,
			compressedSize: entryData.byteLength,
			method: useStore ? 0 : 8,
			crc,
		})
	}

	addStreamingEntry(path: string): void {
		if (this.deflateStream) {
			throw new Error('Streaming entry already open; call closeEntry first')
		}
		this.streamingNameBytes = textEncoder.encode(path)
		this.streamingCrc = 0xffffffff
		this.streamingUncompressedSize = 0
		this.streamingCompressedChunks = []
		this.deflateStream = createDeflateRaw(DEFLATE_OPTS)
		this.deflateStream.on('data', (chunk: Uint8Array) => {
			this.streamingCompressedChunks.push(chunk)
		})
	}

	writeChunk(data: Uint8Array): void {
		if (!this.deflateStream) {
			throw new Error('No streaming entry open; call addStreamingEntry first')
		}
		for (let i = 0; i < data.byteLength; i++) {
			const idx = (this.streamingCrc ^ (data[i] as number)) & 0xff
			this.streamingCrc = ((CRC_TABLE[idx] as number) ^ (this.streamingCrc >>> 8)) >>> 0
		}
		this.streamingUncompressedSize += data.byteLength
		this.deflateStream.write(data)
	}

	closeEntry(): Promise<void> {
		if (!this.deflateStream || !this.streamingNameBytes) {
			throw new Error('No streaming entry open')
		}
		const nameBytes = this.streamingNameBytes
		const stream = this.deflateStream
		if (!nameBytes || !stream) throw new Error('No streaming entry open')
		return new Promise((resolve, reject) => {
			const handleError = (error: Error) => {
				stream.removeListener('end', handleEnd)
				reject(error)
			}
			const handleEnd = () => {
				stream.removeListener('error', handleError)
				const compressedSize = this.streamingCompressedChunks.reduce(
					(sum, c) => sum + c.byteLength,
					0,
				)
				const data = new Uint8Array(compressedSize)
				let offset = 0
				for (const chunk of this.streamingCompressedChunks) {
					data.set(chunk, offset)
					offset += chunk.byteLength
				}
				const crc = (this.streamingCrc ^ 0xffffffff) >>> 0
				this.entries.push({
					nameBytes,
					data,
					uncompressedSize: this.streamingUncompressedSize,
					compressedSize,
					method: 8,
					crc,
				})
				this.streamingNameBytes = null
				this.deflateStream = null
				this.streamingCompressedChunks = []
				resolve()
			}
			stream.once('end', handleEnd)
			stream.once('error', handleError)
			stream.end()
		})
	}

	finalize(): Uint8Array {
		if (this.deflateStream) {
			throw new Error('Streaming entry still open; call closeEntry first')
		}
		return buildZip(this.entries)
	}
}

export function encode(s: string): Uint8Array {
	return textEncoder.encode(s)
}

interface ZipLayoutEntry {
	readonly entry: CompressedEntry
	readonly localOffset: number
	readonly localExtra: Uint8Array
	readonly centralExtra: Uint8Array
	readonly versionNeeded: number
}

function buildZip(entries: readonly CompressedEntry[]): Uint8Array {
	const layout: ZipLayoutEntry[] = []
	let dataSize = 0
	for (const entry of entries) {
		const usesZip64Sizes = entry.compressedSize > UINT32_MAX || entry.uncompressedSize > UINT32_MAX
		const localExtra = usesZip64Sizes
			? buildZip64Extra({
					uncompressedSize: entry.uncompressedSize,
					compressedSize: entry.compressedSize,
				})
			: EMPTY_BYTES
		const localOffset = dataSize
		dataSize += 30 + entry.nameBytes.byteLength + localExtra.byteLength + entry.data.byteLength
		const usesZip64Offset = localOffset > UINT32_MAX
		const centralExtra =
			usesZip64Sizes || usesZip64Offset
				? buildZip64Extra({
						...(usesZip64Sizes ? { uncompressedSize: entry.uncompressedSize } : {}),
						...(usesZip64Sizes ? { compressedSize: entry.compressedSize } : {}),
						...(usesZip64Offset ? { localHeaderOffset: localOffset } : {}),
					})
				: EMPTY_BYTES
		layout.push({
			entry,
			localOffset,
			localExtra,
			centralExtra,
			versionNeeded: usesZip64Sizes || usesZip64Offset ? ZIP64_VERSION : ZIP_VERSION,
		})
	}

	let centralSize = 0
	for (const item of layout) {
		centralSize += 46 + item.entry.nameBytes.byteLength + item.centralExtra.byteLength
	}
	const centralDirOffset = dataSize
	const needsZip64Eocd =
		layout.length > UINT16_MAX ||
		centralDirSizeExceeds32(centralSize) ||
		centralDirOffset > UINT32_MAX
	const out = new Uint8Array(dataSize + centralSize + (needsZip64Eocd ? 56 + 20 : 0) + 22)
	const view = new DataView(out.buffer)
	let offset = 0

	for (const item of layout) {
		const { entry, localExtra, versionNeeded } = item
		view.setUint32(offset, LOCAL_FILE_SIGNATURE, true)
		view.setUint16(offset + 4, versionNeeded, true)
		view.setUint16(offset + 8, entry.method, true)
		view.setUint32(offset + 14, entry.crc, true)
		view.setUint32(
			offset + 18,
			entry.compressedSize > UINT32_MAX ? UINT32_MAX : entry.compressedSize,
			true,
		)
		view.setUint32(
			offset + 22,
			entry.uncompressedSize > UINT32_MAX ? UINT32_MAX : entry.uncompressedSize,
			true,
		)
		view.setUint16(offset + 26, entry.nameBytes.byteLength, true)
		view.setUint16(offset + 28, localExtra.byteLength, true)
		offset += 30
		out.set(entry.nameBytes, offset)
		offset += entry.nameBytes.byteLength
		out.set(localExtra, offset)
		offset += localExtra.byteLength
		out.set(entry.data, offset)
		offset += entry.data.byteLength
	}

	for (const item of layout) {
		const { entry, localOffset, centralExtra, versionNeeded } = item
		view.setUint32(offset, CENTRAL_DIR_SIGNATURE, true)
		view.setUint16(offset + 4, versionNeeded, true)
		view.setUint16(offset + 6, versionNeeded, true)
		view.setUint16(offset + 10, entry.method, true)
		view.setUint32(offset + 16, entry.crc, true)
		view.setUint32(
			offset + 20,
			entry.compressedSize > UINT32_MAX ? UINT32_MAX : entry.compressedSize,
			true,
		)
		view.setUint32(
			offset + 24,
			entry.uncompressedSize > UINT32_MAX ? UINT32_MAX : entry.uncompressedSize,
			true,
		)
		view.setUint16(offset + 28, entry.nameBytes.byteLength, true)
		view.setUint16(offset + 30, centralExtra.byteLength, true)
		view.setUint32(offset + 42, localOffset > UINT32_MAX ? UINT32_MAX : localOffset, true)
		offset += 46
		out.set(entry.nameBytes, offset)
		offset += entry.nameBytes.byteLength
		out.set(centralExtra, offset)
		offset += centralExtra.byteLength
	}

	if (needsZip64Eocd) {
		const zip64EocdOffset = offset
		view.setUint32(offset, ZIP64_EOCD_SIGNATURE, true)
		writeUint64(view, offset + 4, 44)
		view.setUint16(offset + 12, ZIP64_VERSION, true)
		view.setUint16(offset + 14, ZIP64_VERSION, true)
		writeUint64(view, offset + 24, layout.length)
		writeUint64(view, offset + 32, layout.length)
		writeUint64(view, offset + 40, centralSize)
		writeUint64(view, offset + 48, centralDirOffset)
		offset += 56

		view.setUint32(offset, ZIP64_EOCD_LOCATOR_SIGNATURE, true)
		writeUint64(view, offset + 8, zip64EocdOffset)
		view.setUint32(offset + 16, 1, true)
		offset += 20
	}

	view.setUint32(offset, EOCD_SIGNATURE, true)
	view.setUint16(offset + 8, needsZip64Eocd ? UINT16_MAX : layout.length, true)
	view.setUint16(offset + 10, needsZip64Eocd ? UINT16_MAX : layout.length, true)
	view.setUint32(offset + 12, needsZip64Eocd ? UINT32_MAX : centralSize, true)
	view.setUint32(offset + 16, needsZip64Eocd ? UINT32_MAX : centralDirOffset, true)
	return out
}

function centralDirSizeExceeds32(size: number): boolean {
	return size > UINT32_MAX
}

const EMPTY_BYTES = new Uint8Array(0)

function buildZip64Extra(fields: {
	readonly uncompressedSize?: number
	readonly compressedSize?: number
	readonly localHeaderOffset?: number
}): Uint8Array {
	const payloadParts: number[] = []
	if (fields.uncompressedSize !== undefined) payloadParts.push(fields.uncompressedSize)
	if (fields.compressedSize !== undefined) payloadParts.push(fields.compressedSize)
	if (fields.localHeaderOffset !== undefined) payloadParts.push(fields.localHeaderOffset)
	const bytes = new Uint8Array(4 + payloadParts.length * 8)
	const view = new DataView(bytes.buffer)
	view.setUint16(0, ZIP64_EXTRA_FIELD_ID, true)
	view.setUint16(2, payloadParts.length * 8, true)
	let offset = 4
	for (const value of payloadParts) {
		writeUint64(view, offset, value)
		offset += 8
	}
	return bytes
}

function writeUint64(view: DataView, offset: number, value: number): void {
	if (!Number.isFinite(value) || value < 0 || value > Number.MAX_SAFE_INTEGER) {
		throw new Error(`ZIP64 value out of range: ${value}`)
	}
	view.setBigUint64(offset, BigInt(value), true)
}
