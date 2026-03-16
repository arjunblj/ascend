import { createDeflateRaw, deflateRawSync } from 'node:zlib'

const textEncoder = new TextEncoder()
const DEFLATE_OPTS = { level: 2 } as const

const LOCAL_FILE_SIGNATURE = 0x04034b50
const CENTRAL_DIR_SIGNATURE = 0x02014b50
const EOCD_SIGNATURE = 0x06054b50

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
	let dataSize = 0

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
		dataSize += 30 + nameBytes.byteLength + data.byteLength
	}

	let centralSize = 0
	for (const entry of entries) {
		centralSize += 46 + entry.nameBytes.byteLength
	}

	const out = new Uint8Array(dataSize + centralSize + 22)
	const view = new DataView(out.buffer)
	let offset = 0
	const localOffsets: number[] = []

	for (const entry of entries) {
		localOffsets.push(offset)
		view.setUint32(offset, LOCAL_FILE_SIGNATURE, true)
		view.setUint16(offset + 4, 20, true)
		view.setUint16(offset + 8, entry.method, true)
		view.setUint32(offset + 14, entry.crc, true)
		view.setUint32(offset + 18, entry.compressedSize, true)
		view.setUint32(offset + 22, entry.uncompressedSize, true)
		view.setUint16(offset + 26, entry.nameBytes.byteLength, true)
		offset += 30
		out.set(entry.nameBytes, offset)
		offset += entry.nameBytes.byteLength
		out.set(entry.data, offset)
		offset += entry.data.byteLength
	}

	const centralDirOffset = offset
	for (let i = 0; i < entries.length; i++) {
		const entry = entries[i]
		const localOffset = localOffsets[i]
		if (!entry || localOffset === undefined) continue
		view.setUint32(offset, CENTRAL_DIR_SIGNATURE, true)
		view.setUint16(offset + 4, 20, true)
		view.setUint16(offset + 6, 20, true)
		view.setUint16(offset + 10, entry.method, true)
		view.setUint32(offset + 16, entry.crc, true)
		view.setUint32(offset + 20, entry.compressedSize, true)
		view.setUint32(offset + 24, entry.uncompressedSize, true)
		view.setUint16(offset + 28, entry.nameBytes.byteLength, true)
		view.setUint32(offset + 42, localOffset, true)
		offset += 46
		out.set(entry.nameBytes, offset)
		offset += entry.nameBytes.byteLength
	}

	const centralDirSize = offset - centralDirOffset
	view.setUint32(offset, EOCD_SIGNATURE, true)
	view.setUint16(offset + 8, entries.length, true)
	view.setUint16(offset + 10, entries.length, true)
	view.setUint32(offset + 12, centralDirSize, true)
	view.setUint32(offset + 16, centralDirOffset, true)

	return out
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
		let dataSize = 0
		for (const entry of this.entries) {
			dataSize += 30 + entry.nameBytes.byteLength + entry.data.byteLength
		}
		let centralSize = 0
		for (const entry of this.entries) {
			centralSize += 46 + entry.nameBytes.byteLength
		}
		const out = new Uint8Array(dataSize + centralSize + 22)
		const view = new DataView(out.buffer)
		let offset = 0
		const localOffsets: number[] = []

		for (const entry of this.entries) {
			localOffsets.push(offset)
			view.setUint32(offset, LOCAL_FILE_SIGNATURE, true)
			view.setUint16(offset + 4, 20, true)
			view.setUint16(offset + 8, entry.method, true)
			view.setUint32(offset + 14, entry.crc, true)
			view.setUint32(offset + 18, entry.compressedSize, true)
			view.setUint32(offset + 22, entry.uncompressedSize, true)
			view.setUint16(offset + 26, entry.nameBytes.byteLength, true)
			offset += 30
			out.set(entry.nameBytes, offset)
			offset += entry.nameBytes.byteLength
			out.set(entry.data, offset)
			offset += entry.data.byteLength
		}

		const centralDirOffset = offset
		for (let i = 0; i < this.entries.length; i++) {
			const entry = this.entries[i]
			const localOffset = localOffsets[i]
			if (!entry || localOffset === undefined) continue
			view.setUint32(offset, CENTRAL_DIR_SIGNATURE, true)
			view.setUint16(offset + 4, 20, true)
			view.setUint16(offset + 6, 20, true)
			view.setUint16(offset + 10, entry.method, true)
			view.setUint32(offset + 16, entry.crc, true)
			view.setUint32(offset + 20, entry.compressedSize, true)
			view.setUint32(offset + 24, entry.uncompressedSize, true)
			view.setUint16(offset + 28, entry.nameBytes.byteLength, true)
			view.setUint32(offset + 42, localOffset, true)
			offset += 46
			out.set(entry.nameBytes, offset)
			offset += entry.nameBytes.byteLength
		}

		const centralDirSize = offset - centralDirOffset
		view.setUint32(offset, EOCD_SIGNATURE, true)
		view.setUint16(offset + 8, this.entries.length, true)
		view.setUint16(offset + 10, this.entries.length, true)
		view.setUint32(offset + 12, centralDirSize, true)
		view.setUint32(offset + 16, centralDirOffset, true)

		return out
	}
}

export function encode(s: string): Uint8Array {
	return textEncoder.encode(s)
}
