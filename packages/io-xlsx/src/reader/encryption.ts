import { createDecipheriv, createHash, createHmac, timingSafeEqual } from 'node:crypto'
import { type AscendError, ascendError, err, ok, type Result } from '@ascend/schema'
import { isCompoundFile, readCompoundFileStream } from './cfb.ts'

const SEGMENT_LENGTH = 4096
const VERIFIER_HASH_INPUT_BLOCK_KEY = bytes([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79])
const VERIFIER_HASH_VALUE_BLOCK_KEY = bytes([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e])
const ENCRYPTED_KEY_VALUE_BLOCK_KEY = bytes([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6])
const DATA_INTEGRITY_HMAC_KEY_BLOCK_KEY = bytes([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6])
const DATA_INTEGRITY_HMAC_VALUE_BLOCK_KEY = bytes([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33])

interface AgileEncryptionInfo {
	readonly keyData: AgileCipherParams
	readonly encryptedKey: AgilePasswordKey
	readonly encryptedHmacKey?: Uint8Array
	readonly encryptedHmacValue?: Uint8Array
}

interface AgileCipherParams {
	readonly blockSize: number
	readonly keyBits: number
	readonly hashSize: number
	readonly cipherAlgorithm: string
	readonly cipherChaining: string
	readonly hashAlgorithm: string
	readonly saltValue: Uint8Array
}

interface AgilePasswordKey extends AgileCipherParams {
	readonly spinCount: number
	readonly encryptedVerifierHashInput: Uint8Array
	readonly encryptedVerifierHashValue: Uint8Array
	readonly encryptedKeyValue: Uint8Array
}

export function maybeDecryptOoxmlPackage(
	bytes: Uint8Array,
	password: string | undefined,
): Result<Uint8Array | null, AscendError> {
	if (!isCompoundFile(bytes)) return ok(null)
	if (!password) {
		return err(
			ascendError('PROTECTION_ERROR', 'Encrypted XLSX package requires a password', {
				retryStrategy: 'modified',
				details: { encryption: 'agile-ooxml' },
				suggestedFix: 'Pass ReadXlsxOptions.password to decrypt this workbook.',
			}),
		)
	}
	try {
		return ok(decryptAgileOoxmlPackage(bytes, password))
	} catch (error) {
		const message = error instanceof Error ? error.message : 'unknown encryption error'
		if (message === 'invalid-password') {
			return err(
				ascendError('PROTECTION_ERROR', 'Invalid XLSX password', {
					retryStrategy: 'modified',
					details: { encryption: 'agile-ooxml' },
				}),
			)
		}
		if (message.startsWith('Unsupported')) {
			return err(
				ascendError('UNSUPPORTED_FORMAT', message, {
					details: { encryption: 'agile-ooxml' },
				}),
			)
		}
		return err(
			ascendError('CORRUPT_FILE', `Invalid encrypted XLSX package: ${message}`, {
				details: { encryption: 'agile-ooxml' },
			}),
		)
	}
}

function decryptAgileOoxmlPackage(bytes: Uint8Array, password: string): Uint8Array {
	const encryptionInfo = readCompoundFileStream(bytes, 'EncryptionInfo')
	const encryptedPackage = readCompoundFileStream(bytes, 'EncryptedPackage')
	if (!encryptionInfo || !encryptedPackage) {
		throw new Error('Compound file is missing EncryptionInfo or EncryptedPackage')
	}
	const info = parseEncryptionInfo(encryptionInfo)
	validateAgileParams(info.keyData)
	validateAgileParams(info.encryptedKey)
	if (!verifyPassword(password, info.encryptedKey)) throw new Error('invalid-password')
	const secretKey = makeSecretKey(password, info.encryptedKey)
	if (
		info.encryptedHmacKey &&
		info.encryptedHmacValue &&
		!verifyIntegrity(secretKey, info, encryptedPackage)
	) {
		throw new Error('EncryptedPackage integrity check failed')
	}
	return decryptPackagePayload(secretKey, info.keyData, encryptedPackage)
}

function parseEncryptionInfo(bytes: Uint8Array): AgileEncryptionInfo {
	if (bytes.byteLength < 8) throw new Error('EncryptionInfo stream is truncated')
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	const versionMajor = view.getUint16(0, true)
	const versionMinor = view.getUint16(2, true)
	if (versionMajor !== 4 || versionMinor !== 4) {
		throw new Error(`Unsupported EncryptionInfo version ${versionMajor}.${versionMinor}`)
	}
	const xml = new TextDecoder('utf-8').decode(bytes.subarray(8))
	const keyDataAttrs = attrs(matchTag(xml, 'keyData'))
	const dataIntegrityTag = optionalTag(xml, 'dataIntegrity')
	const encryptedKeyAttrs = attrs(matchTag(xml, 'encryptedKey'))
	return {
		keyData: parseCipherParams(keyDataAttrs),
		encryptedKey: {
			...parseCipherParams(encryptedKeyAttrs),
			spinCount: requiredInt(encryptedKeyAttrs, 'spinCount'),
			encryptedVerifierHashInput: requiredBase64(encryptedKeyAttrs, 'encryptedVerifierHashInput'),
			encryptedVerifierHashValue: requiredBase64(encryptedKeyAttrs, 'encryptedVerifierHashValue'),
			encryptedKeyValue: requiredBase64(encryptedKeyAttrs, 'encryptedKeyValue'),
		},
		...(dataIntegrityTag
			? {
					encryptedHmacKey: requiredBase64(attrs(dataIntegrityTag), 'encryptedHmacKey'),
					encryptedHmacValue: requiredBase64(attrs(dataIntegrityTag), 'encryptedHmacValue'),
				}
			: {}),
	}
}

function parseCipherParams(attributes: Record<string, string>): AgileCipherParams {
	return {
		blockSize: requiredInt(attributes, 'blockSize'),
		keyBits: requiredInt(attributes, 'keyBits'),
		hashSize: requiredInt(attributes, 'hashSize'),
		cipherAlgorithm: required(attributes, 'cipherAlgorithm'),
		cipherChaining: required(attributes, 'cipherChaining'),
		hashAlgorithm: required(attributes, 'hashAlgorithm'),
		saltValue: requiredBase64(attributes, 'saltValue'),
	}
}

function validateAgileParams(params: AgileCipherParams): void {
	if (params.cipherAlgorithm !== 'AES') {
		throw new Error(`Unsupported agile cipher ${params.cipherAlgorithm}`)
	}
	if (params.cipherChaining !== 'ChainingModeCBC') {
		throw new Error(`Unsupported agile cipher chaining ${params.cipherChaining}`)
	}
	if (params.keyBits !== 128 && params.keyBits !== 192 && params.keyBits !== 256) {
		throw new Error(`Unsupported agile AES key size ${params.keyBits}`)
	}
	if (params.blockSize !== 16) throw new Error(`Unsupported agile block size ${params.blockSize}`)
	hashName(params.hashAlgorithm)
}

function verifyPassword(password: string, key: AgilePasswordKey): boolean {
	const hash = deriveIteratedHash(password, key)
	const verifierKey = deriveEncryptionKey(hash, VERIFIER_HASH_INPUT_BLOCK_KEY, key)
	const verifierHashKey = deriveEncryptionKey(hash, VERIFIER_HASH_VALUE_BLOCK_KEY, key)
	const verifierInput = decryptAesCbc(key.encryptedVerifierHashInput, verifierKey, key.saltValue)
	const actualHash = digest(key.hashAlgorithm, verifierInput)
	const expectedHash = decryptAesCbc(key.encryptedVerifierHashValue, verifierHashKey, key.saltValue)
	return safeEqual(actualHash, expectedHash.subarray(0, actualHash.byteLength))
}

function makeSecretKey(password: string, key: AgilePasswordKey): Uint8Array {
	const hash = deriveIteratedHash(password, key)
	const encryptionKey = deriveEncryptionKey(hash, ENCRYPTED_KEY_VALUE_BLOCK_KEY, key)
	return decryptAesCbc(key.encryptedKeyValue, encryptionKey, key.saltValue)
}

function verifyIntegrity(
	secretKey: Uint8Array,
	info: AgileEncryptionInfo,
	encryptedPackage: Uint8Array,
): boolean {
	if (!info.encryptedHmacKey || !info.encryptedHmacValue) return true
	const hmacKeyIv = blockIv(info.keyData, DATA_INTEGRITY_HMAC_KEY_BLOCK_KEY)
	const hmacValueIv = blockIv(info.keyData, DATA_INTEGRITY_HMAC_VALUE_BLOCK_KEY)
	const hmacKey = decryptAesCbc(info.encryptedHmacKey, secretKey, hmacKeyIv)
	const hmacValue = decryptAesCbc(info.encryptedHmacValue, secretKey, hmacValueIv)
	const actual = Uint8Array.from(
		createHmac(hashName(info.keyData.hashAlgorithm), hmacKey).update(encryptedPackage).digest(),
	)
	return safeEqual(actual, hmacValue.subarray(0, actual.byteLength))
}

function decryptPackagePayload(
	secretKey: Uint8Array,
	params: AgileCipherParams,
	encryptedPackage: Uint8Array,
): Uint8Array {
	if (encryptedPackage.byteLength < 8) throw new Error('EncryptedPackage stream is truncated')
	const view = new DataView(
		encryptedPackage.buffer,
		encryptedPackage.byteOffset,
		encryptedPackage.byteLength,
	)
	const totalSize = readU64(view, 0)
	const output = new Uint8Array(totalSize)
	let encryptedOffset = 8
	let outputOffset = 0
	for (
		let segment = 0;
		outputOffset < totalSize && encryptedOffset < encryptedPackage.byteLength;
		segment++
	) {
		const chunk = encryptedPackage.subarray(
			encryptedOffset,
			Math.min(encryptedOffset + SEGMENT_LENGTH, encryptedPackage.byteLength),
		)
		if (chunk.byteLength % params.blockSize !== 0) {
			throw new Error('EncryptedPackage segment is not block aligned')
		}
		const decrypted = decryptAesCbc(chunk, secretKey, segmentIv(params, segment))
		const take = Math.min(decrypted.byteLength, totalSize - outputOffset)
		output.set(decrypted.subarray(0, take), outputOffset)
		outputOffset += take
		encryptedOffset += chunk.byteLength
	}
	if (outputOffset !== totalSize) throw new Error('EncryptedPackage payload is truncated')
	return output
}

function deriveIteratedHash(password: string, key: AgilePasswordKey): Uint8Array {
	let hash = digest(key.hashAlgorithm, concat([key.saltValue, utf16le(password)]))
	const iterator = new Uint8Array(4)
	const iteratorView = new DataView(iterator.buffer)
	for (let i = 0; i < key.spinCount; i++) {
		iteratorView.setUint32(0, i, true)
		hash = digest(key.hashAlgorithm, concat([iterator, hash]))
	}
	return hash
}

function deriveEncryptionKey(
	iteratedHash: Uint8Array,
	blockKey: Uint8Array,
	params: AgileCipherParams,
): Uint8Array {
	return digest(params.hashAlgorithm, concat([iteratedHash, blockKey])).subarray(
		0,
		params.keyBits / 8,
	)
}

function segmentIv(params: AgileCipherParams, segment: number): Uint8Array {
	const blockKey = new Uint8Array(4)
	new DataView(blockKey.buffer).setUint32(0, segment, true)
	return blockIv(params, blockKey)
}

function blockIv(params: AgileCipherParams, blockKey: Uint8Array): Uint8Array {
	return digest(params.hashAlgorithm, concat([params.saltValue, blockKey])).subarray(
		0,
		params.blockSize,
	)
}

function decryptAesCbc(data: Uint8Array, key: Uint8Array, iv: Uint8Array): Uint8Array {
	const decipher = createDecipheriv(`aes-${key.byteLength * 8}-cbc`, key, iv)
	decipher.setAutoPadding(false)
	return concat([Uint8Array.from(decipher.update(data)), Uint8Array.from(decipher.final())])
}

function digest(algorithm: string, data: Uint8Array): Uint8Array {
	return Uint8Array.from(createHash(hashName(algorithm)).update(data).digest())
}

function hashName(algorithm: string): string {
	switch (algorithm.toUpperCase().replaceAll('-', '')) {
		case 'SHA1':
			return 'sha1'
		case 'SHA256':
			return 'sha256'
		case 'SHA384':
			return 'sha384'
		case 'SHA512':
			return 'sha512'
		default:
			throw new Error(`Unsupported agile hash algorithm ${algorithm}`)
	}
}

function matchTag(xml: string, tagName: string): string {
	const tag = optionalTag(xml, tagName)
	if (!tag) throw new Error(`EncryptionInfo missing ${tagName}`)
	return tag
}

function optionalTag(xml: string, tagName: string): string | undefined {
	return xml.match(new RegExp(`<(?:\\w+:)?${tagName}\\b[^>]*>`, 'u'))?.[0]
}

function attrs(tag: string): Record<string, string> {
	const out: Record<string, string> = {}
	for (const match of tag.matchAll(/([\w:]+)="([^"]*)"/gu)) {
		out[match[1] ?? ''] = match[2] ?? ''
	}
	return out
}

function required(attributes: Record<string, string>, name: string): string {
	const value = attributes[name]
	if (value === undefined || value === '') throw new Error(`EncryptionInfo missing ${name}`)
	return value
}

function requiredInt(attributes: Record<string, string>, name: string): number {
	const value = Number.parseInt(required(attributes, name), 10)
	if (!Number.isFinite(value)) throw new Error(`EncryptionInfo has invalid ${name}`)
	return value
}

function requiredBase64(attributes: Record<string, string>, name: string): Uint8Array {
	return Uint8Array.from(Buffer.from(required(attributes, name), 'base64'))
}

function safeEqual(a: Uint8Array, b: Uint8Array): boolean {
	return a.byteLength === b.byteLength && timingSafeEqual(a, b)
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

function bytes(values: readonly number[]): Uint8Array {
	return Uint8Array.from(values)
}

function utf16le(text: string): Uint8Array {
	const out = new Uint8Array(text.length * 2)
	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i)
		out[i * 2] = code & 0xff
		out[i * 2 + 1] = code >>> 8
	}
	return out
}

function readU64(view: DataView, offset: number): number {
	const low = BigInt(view.getUint32(offset, true))
	const high = BigInt(view.getUint32(offset + 4, true))
	const value = (high << 32n) + low
	if (value > BigInt(Number.MAX_SAFE_INTEGER))
		throw new Error('EncryptedPackage payload is too large')
	return Number(value)
}
