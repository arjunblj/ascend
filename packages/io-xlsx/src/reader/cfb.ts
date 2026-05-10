const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const
const END_OF_CHAIN = 0xfffffffe
const FAT_SECTOR = 0xfffffffd
const HEADER_DIFAT_ENTRIES = 109
const MINI_STREAM_CUTOFF = 4096
const DIRECTORY_ENTRY_BYTES = 128

interface DirectoryEntry {
	readonly name: string
	readonly type: number
	readonly startSector: number
	readonly size: number
}

export function isCompoundFile(bytes: Uint8Array): boolean {
	return CFB_MAGIC.every((byte, index) => bytes[index] === byte)
}

export function readCompoundFileStream(
	bytes: Uint8Array,
	streamName: string,
): Uint8Array | undefined {
	const container = parseCompoundFile(bytes)
	return container.readStream(streamName)
}

function parseCompoundFile(bytes: Uint8Array): {
	readStream(streamName: string): Uint8Array | undefined
} {
	if (!isCompoundFile(bytes)) throw new Error('Not a compound file')
	if (bytes.byteLength < 512) throw new Error('Compound file header is truncated')

	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const sectorSize = 1 << readU16(view, 30)
	const miniSectorSize = 1 << readU16(view, 32)
	const directoryStart = readU32(view, 48)
	const miniStreamCutoff = readU32(view, 56) || MINI_STREAM_CUTOFF
	const miniFatStart = readU32(view, 60)
	const miniFatSectorCount = readU32(view, 64)
	const difatStart = readU32(view, 68)
	const difatSectorCount = readU32(view, 72)
	const difat = readDifat(bytes, view, sectorSize, difatStart, difatSectorCount)
	const fat = readFat(bytes, sectorSize, difat)

	const readRegularChain = (startSector: number, maxBytes?: number): Uint8Array => {
		if (startSector >= END_OF_CHAIN) return new Uint8Array()
		const chunks: Uint8Array[] = []
		const seen = new Set<number>()
		let sector = startSector
		let remaining = maxBytes ?? Number.POSITIVE_INFINITY
		while (sector < END_OF_CHAIN && remaining > 0) {
			if (seen.has(sector)) throw new Error('Compound file sector chain loop')
			seen.add(sector)
			const chunk = readSector(bytes, sectorSize, sector)
			const take = Math.min(chunk.byteLength, remaining)
			chunks.push(chunk.subarray(0, take))
			remaining -= take
			sector = fat[sector] ?? END_OF_CHAIN
		}
		return concat(chunks)
	}

	const directoryBytes = readRegularChain(directoryStart)
	const directory = parseDirectory(directoryBytes)
	const root = directory.find((entry) => entry.type === 5)
	const miniFat =
		miniFatStart < END_OF_CHAIN && miniFatSectorCount > 0
			? readMiniFat(readRegularChain(miniFatStart, miniFatSectorCount * sectorSize))
			: []
	const miniStream =
		root && root.startSector < END_OF_CHAIN
			? readRegularChain(root.startSector, root.size)
			: new Uint8Array()

	const readMiniChain = (startSector: number, maxBytes: number): Uint8Array => {
		const chunks: Uint8Array[] = []
		const seen = new Set<number>()
		let sector = startSector
		let remaining = maxBytes
		while (sector < END_OF_CHAIN && remaining > 0) {
			if (seen.has(sector)) throw new Error('Compound file mini sector chain loop')
			seen.add(sector)
			const offset = sector * miniSectorSize
			if (offset >= miniStream.byteLength) break
			const chunk = miniStream.subarray(
				offset,
				Math.min(offset + miniSectorSize, miniStream.byteLength),
			)
			const take = Math.min(chunk.byteLength, remaining)
			chunks.push(chunk.subarray(0, take))
			remaining -= take
			sector = miniFat[sector] ?? END_OF_CHAIN
		}
		return concat(chunks)
	}

	return {
		readStream(streamName: string): Uint8Array | undefined {
			const normalized = streamName.replace(/^\//, '')
			const entry = directory.find((item) => item.type === 2 && item.name === normalized)
			if (!entry) return undefined
			if (entry.size < miniStreamCutoff) return readMiniChain(entry.startSector, entry.size)
			return readRegularChain(entry.startSector, entry.size)
		},
	}
}

function readDifat(
	bytes: Uint8Array,
	view: DataView,
	sectorSize: number,
	difatStart: number,
	difatSectorCount: number,
): number[] {
	const difat: number[] = []
	for (let offset = 76; offset < 76 + HEADER_DIFAT_ENTRIES * 4; offset += 4) {
		const sector = readU32(view, offset)
		if (sector < FAT_SECTOR) difat.push(sector)
	}
	let next = difatStart
	for (let i = 0; i < difatSectorCount && next < END_OF_CHAIN; i++) {
		const sector = readSector(bytes, sectorSize, next)
		const sectorView = dataView(sector)
		for (let offset = 0; offset < sectorSize - 4; offset += 4) {
			const fatSector = readU32(sectorView, offset)
			if (fatSector < FAT_SECTOR) difat.push(fatSector)
		}
		next = readU32(sectorView, sectorSize - 4)
	}
	return difat
}

function readFat(bytes: Uint8Array, sectorSize: number, difat: readonly number[]): number[] {
	const fat: number[] = []
	for (const sectorId of difat) {
		const sector = readSector(bytes, sectorSize, sectorId)
		const sectorView = dataView(sector)
		for (let offset = 0; offset < sectorSize; offset += 4) {
			fat.push(readU32(sectorView, offset))
		}
	}
	return fat
}

function readMiniFat(bytes: Uint8Array): number[] {
	const view = dataView(bytes)
	const fat: number[] = []
	for (let offset = 0; offset + 4 <= bytes.byteLength; offset += 4) {
		fat.push(readU32(view, offset))
	}
	return fat
}

function parseDirectory(bytes: Uint8Array): DirectoryEntry[] {
	const entries: DirectoryEntry[] = []
	const view = dataView(bytes)
	for (
		let offset = 0;
		offset + DIRECTORY_ENTRY_BYTES <= bytes.byteLength;
		offset += DIRECTORY_ENTRY_BYTES
	) {
		const nameLength = readU16(view, offset + 64)
		if (nameLength < 2) continue
		const nameBytes = bytes.subarray(offset, offset + nameLength - 2)
		const name = new TextDecoder('utf-16le').decode(nameBytes)
		const type = bytes[offset + 66] ?? 0
		if (type === 0) continue
		entries.push({
			name,
			type,
			startSector: readU32(view, offset + 116),
			size: readU64(view, offset + 120),
		})
	}
	return entries
}

function readSector(bytes: Uint8Array, sectorSize: number, sectorId: number): Uint8Array {
	if (sectorId >= FAT_SECTOR) return new Uint8Array()
	const offset = (sectorId + 1) * sectorSize
	if (offset + sectorSize > bytes.byteLength) {
		throw new Error(`Compound file sector ${sectorId} is out of range`)
	}
	return bytes.subarray(offset, offset + sectorSize)
}

function concat(chunks: readonly Uint8Array[]): Uint8Array {
	const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0)
	const out = new Uint8Array(total)
	let offset = 0
	for (const chunk of chunks) {
		out.set(chunk, offset)
		offset += chunk.byteLength
	}
	return out
}

function dataView(bytes: Uint8Array): DataView {
	return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function readU16(view: DataView, offset: number): number {
	return view.getUint16(offset, true)
}

function readU32(view: DataView, offset: number): number {
	return view.getUint32(offset, true)
}

function readU64(view: DataView, offset: number): number {
	const low = BigInt(view.getUint32(offset, true))
	const high = BigInt(view.getUint32(offset + 4, true))
	const value = (high << 32n) + low
	if (value > BigInt(Number.MAX_SAFE_INTEGER)) throw new Error('Compound file stream is too large')
	return Number(value)
}
