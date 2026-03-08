import { unzipSync } from 'fflate'

export function extractZip(bytes: Uint8Array): Map<string, Uint8Array> {
	const entries = unzipSync(bytes)
	const parts = new Map<string, Uint8Array>()
	for (const [path, data] of Object.entries(entries)) {
		parts.set(path, data)
	}
	return parts
}
