import { strToU8, zipSync } from 'fflate'

export function makeXlsx(parts: Record<string, string>): Uint8Array {
	const entries: Record<string, Uint8Array> = {}
	for (const [path, content] of Object.entries(parts)) {
		entries[path] = strToU8(content)
	}
	return zipSync(entries)
}
