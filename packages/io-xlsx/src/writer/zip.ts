import { strToU8, zipSync } from 'fflate'

export function createZip(parts: Map<string, Uint8Array>): Uint8Array {
	const entries: Record<string, Uint8Array> = {}
	for (const [path, data] of parts) {
		entries[path] = data
	}
	return zipSync(entries)
}

export function encode(s: string): Uint8Array {
	return strToU8(s)
}
