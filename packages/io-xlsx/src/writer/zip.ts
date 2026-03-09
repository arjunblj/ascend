import { strToU8, zipSync } from 'fflate'

const COMPRESS_OPTS = { level: 2 as const }

export function createZip(parts: Map<string, Uint8Array>): Uint8Array {
	const entries: Record<string, [Uint8Array, typeof COMPRESS_OPTS]> = {}
	for (const [path, data] of parts) {
		entries[path] = [data, COMPRESS_OPTS]
	}
	return zipSync(entries)
}

export function encode(s: string): Uint8Array {
	return strToU8(s)
}
