import { describe, expect, test } from 'bun:test'
import { createZip, encode } from '../writer/zip.ts'
import { extractZip } from './zip.ts'

const EOCD_SIGNATURE = 0x06054b50
const ZIP64_EOCD_SIGNATURE = 0x06064b50
const ZIP64_EOCD_LOCATOR_SIGNATURE = 0x07064b50
const ZIP64_EXTRA_FIELD_ID = 0x0001
const LOCAL_FILE_SIGNATURE = 0x04034b50
const CENTRAL_DIR_SIGNATURE = 0x02014b50

describe('extractZip', () => {
	test('writer records standard CRC32 values in ZIP headers', () => {
		const bytes = createZip(new Map([['hello.txt', encode('hello')]]))
		const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
		const localOffset = findSignature(bytes, LOCAL_FILE_SIGNATURE)
		const centralOffset = findSignature(bytes, CENTRAL_DIR_SIGNATURE)
		expect(view.getUint32(localOffset + 14, true)).toBe(0x3610a686)
		expect(view.getUint32(centralOffset + 16, true)).toBe(0x3610a686)
	})

	test('reads ZIP64 archives when central directory uses ZIP64 metadata', () => {
		const base = createZip(new Map([['hello.txt', encode('hello zip64')]]))
		const zip64Bytes = promoteToZip64(base)
		const archive = extractZip(zip64Bytes)
		expect(archive.readText('hello.txt')).toBe('hello zip64')
	})

	test('streams deflated text chunks without changing decoded content', () => {
		const text = 'alpha😀beta\n'.repeat(256)
		const archive = extractZip(createZip(new Map([['hello.txt', encode(text)]])))
		expect([...archive.readTextChunks('hello.txt', 7)].join('')).toBe(text)
	})

	test('streams deflated byte chunks without changing bytes', async () => {
		const text = 'alpha😀beta\n'.repeat(256)
		const expected = encode(text)
		const archive = extractZip(createZip(new Map([['hello.txt', expected]])))
		const chunks: Uint8Array[] = []
		for await (const chunk of archive.readByteChunksAsync('hello.txt', 11)) chunks.push(chunk)
		expect(concatBytes(chunks)).toEqual(expected)
	})

	test('can prefer streaming byte chunks without caching partial output', () => {
		const text = 'alpha😀beta\n'.repeat(4096)
		const expected = encode(text)
		const archive = extractZip(createZip(new Map([['hello.txt', expected]])))
		const iterator = archive.readByteChunks('hello.txt', 128, { preferStreaming: true })
		const first = iterator.next()
		expect(first.done).toBe(false)
		expect(first.value.byteLength).toBeGreaterThan(0)
		iterator.return?.()
		expect(archive.readBytes('hello.txt')).toEqual(expected)
	})

	test('streams stored text chunks without changing decoded content', () => {
		const text = 'x'
		const archive = extractZip(createZip(new Map([['stored.txt', encode(text)]])))
		expect(archive.get('stored.txt')?.compressionMethod).toBe(0)
		expect([...archive.readTextChunks('stored.txt', 1)].join('')).toBe(text)
	})

	test('reads ZIP64 archives emitted by writer when entry count exceeds EOCD16 limits', () => {
		const parts = new Map<string, Uint8Array>()
		for (let i = 0; i < 70_000; i++) {
			parts.set(`f${i}.txt`, new Uint8Array(0))
		}
		const bytes = createZip(parts)
		const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
		const eocdOffset = findSignature(bytes, EOCD_SIGNATURE)
		expect(view.getUint16(eocdOffset + 8, true)).toBe(0xffff)
		expect(view.getUint16(eocdOffset + 10, true)).toBe(0xffff)
		expect(findSignature(bytes, ZIP64_EOCD_SIGNATURE)).toBeGreaterThan(0)
		expect(findSignature(bytes, ZIP64_EOCD_LOCATOR_SIGNATURE)).toBeGreaterThan(0)

		const archive = extractZip(bytes)
		expect(archive.readBytes('f0.txt')?.byteLength).toBe(0)
		expect(archive.readBytes('f69999.txt')?.byteLength).toBe(0)
		let entryCount = 0
		for (const _ of archive.entries()) entryCount++
		expect(entryCount).toBe(70_000)
	})
})

function promoteToZip64(bytes: Uint8Array): Uint8Array {
	const eocdOffset = findSignature(bytes, EOCD_SIGNATURE)
	const baseView = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const entryCount = baseView.getUint16(eocdOffset + 10, true)
	const centralDirSize = baseView.getUint32(eocdOffset + 12, true)
	const centralDirOffset = baseView.getUint32(eocdOffset + 16, true)
	const centralEntryOffset = centralDirOffset
	if (baseView.getUint32(centralEntryOffset, true) !== CENTRAL_DIR_SIGNATURE) {
		throw new Error('Expected a central directory entry')
	}

	const fileNameLength = baseView.getUint16(centralEntryOffset + 28, true)
	const extraFieldLength = baseView.getUint16(centralEntryOffset + 30, true)
	const fileCommentLength = baseView.getUint16(centralEntryOffset + 32, true)
	const compressedSize = baseView.getUint32(centralEntryOffset + 20, true)
	const uncompressedSize = baseView.getUint32(centralEntryOffset + 24, true)
	const fileNameStart = centralEntryOffset + 46
	const fileNameEnd = fileNameStart + fileNameLength
	const extraFieldStart = fileNameEnd
	const extraFieldEnd = extraFieldStart + extraFieldLength
	const centralEntryEnd = extraFieldEnd + fileCommentLength

	const zip64Extra = buildZip64Extra(uncompressedSize, compressedSize)
	const centralDirWithZip64 = concatBytes([
		bytes.subarray(centralDirOffset, extraFieldStart),
		zip64Extra,
		bytes.subarray(extraFieldStart, centralEntryEnd),
	])
	const centralView = new DataView(
		centralDirWithZip64.buffer,
		centralDirWithZip64.byteOffset,
		centralDirWithZip64.byteLength,
	)
	centralView.setUint32(20, 0xffffffff, true)
	centralView.setUint32(24, 0xffffffff, true)
	centralView.setUint16(30, extraFieldLength + zip64Extra.byteLength, true)

	const patchedCentralDirSize = centralDirSize + zip64Extra.byteLength
	const zip64Eocd = buildZip64Eocd(entryCount, patchedCentralDirSize, centralDirOffset)
	const zip64Locator = buildZip64Locator(centralDirOffset + patchedCentralDirSize)
	const eocd = bytes.slice(eocdOffset)
	const eocdView = new DataView(eocd.buffer, eocd.byteOffset, eocd.byteLength)
	eocdView.setUint16(8, 0xffff, true)
	eocdView.setUint16(10, 0xffff, true)
	eocdView.setUint32(12, 0xffffffff, true)
	eocdView.setUint32(16, 0xffffffff, true)

	return concatBytes([
		bytes.subarray(0, centralDirOffset),
		centralDirWithZip64,
		zip64Eocd,
		zip64Locator,
		eocd,
	])
}

function buildZip64Extra(uncompressedSize: number, compressedSize: number): Uint8Array {
	const bytes = new Uint8Array(4 + 16)
	const view = new DataView(bytes.buffer)
	view.setUint16(0, ZIP64_EXTRA_FIELD_ID, true)
	view.setUint16(2, 16, true)
	view.setBigUint64(4, BigInt(uncompressedSize), true)
	view.setBigUint64(12, BigInt(compressedSize), true)
	return bytes
}

function buildZip64Eocd(
	entryCount: number,
	centralDirSize: number,
	centralDirOffset: number,
): Uint8Array {
	const bytes = new Uint8Array(56)
	const view = new DataView(bytes.buffer)
	view.setUint32(0, ZIP64_EOCD_SIGNATURE, true)
	view.setBigUint64(4, BigInt(44), true)
	view.setUint16(12, 45, true)
	view.setUint16(14, 45, true)
	view.setUint32(16, 0, true)
	view.setUint32(20, 0, true)
	view.setBigUint64(24, BigInt(entryCount), true)
	view.setBigUint64(32, BigInt(entryCount), true)
	view.setBigUint64(40, BigInt(centralDirSize), true)
	view.setBigUint64(48, BigInt(centralDirOffset), true)
	return bytes
}

function buildZip64Locator(zip64EocdOffset: number): Uint8Array {
	const bytes = new Uint8Array(20)
	const view = new DataView(bytes.buffer)
	view.setUint32(0, ZIP64_EOCD_LOCATOR_SIGNATURE, true)
	view.setUint32(4, 0, true)
	view.setBigUint64(8, BigInt(zip64EocdOffset), true)
	view.setUint32(16, 1, true)
	return bytes
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
	const total = parts.reduce((sum, part) => sum + part.byteLength, 0)
	const out = new Uint8Array(total)
	let offset = 0
	for (const part of parts) {
		out.set(part, offset)
		offset += part.byteLength
	}
	return out
}

function findSignature(bytes: Uint8Array, signature: number): number {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	for (let offset = bytes.byteLength - 4; offset >= 0; offset--) {
		if (view.getUint32(offset, true) === signature) return offset
	}
	throw new Error(`Missing signature ${signature.toString(16)}`)
}
