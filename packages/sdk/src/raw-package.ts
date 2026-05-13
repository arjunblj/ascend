import { Buffer } from 'node:buffer'
import { createHash } from 'node:crypto'
import { extractZip, inspectXlsxPackageGraph } from '@ascend/io-xlsx'
import type { RawPackagePartInfo, RawPackagePartOptions } from './types.ts'

const DEFAULT_MAX_RAW_PART_BYTES = 256 * 1024
const RAW_PACKAGE_SEMANTICS = 'raw-package-bytes' as const

export function inspectRawPackagePart(
	bytes: Uint8Array,
	options: RawPackagePartOptions,
): RawPackagePartInfo {
	const requestedPartPath = options.partPath
	const normalized = normalizePackagePartPath(requestedPartPath)
	const baseInfo = {
		requestedPartPath,
		partPath: normalized.ok ? normalized.path : requestedPartPath,
		found: false,
		validPath: normalized.ok,
		semantics: RAW_PACKAGE_SEMANTICS,
		...(options.origin ? { origin: options.origin } : {}),
		...(normalized.ok && normalized.normalizedFromRoot
			? { normalizedFromRoot: normalized.normalizedFromRoot }
			: {}),
		caseInsensitiveRequested: options.caseInsensitive === true,
		caseInsensitiveFallback: false,
	}
	if (!normalized.ok) {
		return {
			...baseInfo,
			invalidReason: normalized.reason,
		}
	}
	const archive = extractZip(bytes)
	const exact = archive.has(normalized.path)
	const fallbackPath =
		!exact && options.caseInsensitive === true
			? archive.resolvePathCaseInsensitive(normalized.path)
			: null
	const partPath = fallbackPath ?? normalized.path
	const partBytes = archive.readBytes(partPath)
	if (!partBytes) {
		return {
			...baseInfo,
			partPath,
		}
	}
	const graph = inspectXlsxPackageGraph(bytes)
	const part = graph.parts.find((entry) => entry.path === partPath)
	const maxBytes = Math.max(0, options.maxBytes ?? DEFAULT_MAX_RAW_PART_BYTES)
	const encoding = options.encoding ?? 'text'
	const previewByteLength = encoding === 'none' ? 0 : Math.min(maxBytes, partBytes.byteLength)
	const truncated = encoding !== 'none' && partBytes.byteLength > previewByteLength
	const previewBytes = partBytes.subarray(0, previewByteLength)
	const binaryLike = looksBinary(partBytes)
	return {
		...baseInfo,
		partPath,
		found: true,
		caseInsensitiveFallback: fallbackPath !== null,
		...(part
			? {
					contentType: part.contentType,
					contentTypeSource: part.contentTypeSource,
					featureFamily: part.featureFamily,
					ownerScope: part.ownerScope,
					preservationPolicy: part.preservationPolicy,
					bytePreservationExpected: part.bytePreservationExpected,
				}
			: {}),
		byteLength: partBytes.byteLength,
		sha256: createHash('sha256').update(partBytes).digest('hex'),
		encoding,
		previewByteLength,
		truncated,
		maxBytes,
		binaryLike,
		...(encoding === 'text' && binaryLike
			? {
					textWarning:
						'Part appears binary; text is a diagnostic UTF-8 preview and not semantic workbook truth.',
				}
			: {}),
		...(encoding === 'text' ? { text: new TextDecoder('utf-8').decode(previewBytes) } : {}),
		...(encoding === 'base64' ? { base64: Buffer.from(previewBytes).toString('base64') } : {}),
	}
}

function normalizePackagePartPath(
	path: string,
): { ok: true; path: string; normalizedFromRoot?: boolean } | { ok: false; reason: string } {
	const trimmed = path.trim()
	if (trimmed.length === 0) return { ok: false, reason: 'Package part path is empty.' }
	if (trimmed.includes('\\')) {
		return { ok: false, reason: 'Package part path must use forward slashes.' }
	}
	if (trimmed.startsWith('//')) {
		return { ok: false, reason: 'Package part path must not contain duplicate slashes.' }
	}
	const normalizedFromRoot = trimmed.startsWith('/')
	const normalized = normalizedFromRoot ? trimmed.slice(1) : trimmed
	if (normalized.length === 0) return { ok: false, reason: 'Package part path is empty.' }
	if (normalized.includes('//')) {
		return { ok: false, reason: 'Package part path must not contain duplicate slashes.' }
	}
	if (normalized.split('/').some((segment) => segment === '..' || segment === '.')) {
		return { ok: false, reason: 'Package part path must not contain dot segments.' }
	}
	return { ok: true, path: normalized, ...(normalizedFromRoot ? { normalizedFromRoot } : {}) }
}

function looksBinary(bytes: Uint8Array): boolean {
	const sample = bytes.subarray(0, Math.min(bytes.byteLength, 4096))
	if (sample.length === 0) return false
	let suspicious = 0
	for (const byte of sample) {
		if (byte === 0) return true
		if (byte < 0x09 || (byte > 0x0d && byte < 0x20)) suspicious += 1
	}
	return suspicious / sample.length > 0.05
}
