import { Buffer } from 'node:buffer'

export interface WasmRangeOps {
	load(values: Float64Array, count: number): void
	/** Writable view into linear WASM memory for direct numeric fills (avoids an extra JS copy before sum/min/max). */
	numericScratch(minDoubles: number): Float64Array | null
	sum(count: number): number
	min(count: number): number
	max(count: number): number
}

const WASM_BASE64 =
	'AGFzbQEAAAABBwFgAn9/AXwDBAMAAAAFAwEAAQcoBAZtZW1vcnkCAAdzdW1fZjY0AAAHbWluX2Y2NAABB21heF9mNjQAAgquAwN2AwF/AXsBfEQAAAAAAAAAAP0UIQMCQANAIAJBAWogAU8NASADIAAgAkEDdGr9AAQA/fABIQMgAkECaiECDAALCyAD/SEAIAP9IQGgIQQCQANAIAIgAU8NASAEIAAgAkEDdGorAwCgIQQgAkEBaiECDAALCyAEC5kBAwF/AXsBfCABRQR8RAAAAAAAAAAABSABQQFLBEAgAP0ABAAhA0ECIQICQANAIAJBAWogAU8NASADIAAgAkEDdGr9AAQA/fQBIQMgAkECaiECDAALCyAD/SEAIAP9IQGkIQQFIAArAwAhBEEBIQILAkADQCACIAFPDQEgBCAAIAJBA3RqKwMApCEEIAJBAWohAgwACwsgBAsLmQEDAX8BewF8IAFFBHxEAAAAAAAAAAAFIAFBAUsEQCAA/QAEACEDQQIhAgJAA0AgAkEBaiABTw0BIAMgACACQQN0av0ABAD99QEhAyACQQJqIQIMAAsLIAP9IQAgA/0hAaUhBAUgACsDACEEQQEhAgsCQANAIAIgAU8NASAEIAAgAkEDdGorAwClIQQgAkEBaiECDAALCyAECws='

function initWasmRangeOps(): WasmRangeOps | null {
	try {
		const module = new WebAssembly.Module(Buffer.from(WASM_BASE64, 'base64'))
		const instance = new WebAssembly.Instance(module, {})
		const exports = instance.exports as {
			memory: WebAssembly.Memory
			sum_f64(ptr: number, len: number): number
			min_f64(ptr: number, len: number): number
			max_f64(ptr: number, len: number): number
		}
		let view = new Float64Array(exports.memory.buffer)

		function ensureCapacity(count: number): void {
			const requiredBytes = Math.max(8, count * 8)
			const currentBytes = exports.memory.buffer.byteLength
			if (requiredBytes > currentBytes) {
				const currentPages = currentBytes >>> 16
				const requiredPages = Math.ceil(requiredBytes / 65_536)
				exports.memory.grow(requiredPages - currentPages)
			}
			if (view.buffer !== exports.memory.buffer) {
				view = new Float64Array(exports.memory.buffer)
			}
		}

		return {
			load(values: Float64Array, count: number): void {
				ensureCapacity(count)
				view.set(values.subarray(0, count), 0)
			},
			numericScratch(minDoubles: number): Float64Array | null {
				if (minDoubles < 128) return null
				ensureCapacity(minDoubles)
				return view.subarray(0, minDoubles)
			},
			sum(count: number): number {
				return exports.sum_f64(0, count)
			},
			min(count: number): number {
				return exports.min_f64(0, count)
			},
			max(count: number): number {
				return exports.max_f64(0, count)
			},
		}
	} catch {
		return null
	}
}

let wasmRangeOps: WasmRangeOps | null | undefined

export function getWasmRangeOps(): WasmRangeOps | null {
	if (wasmRangeOps === undefined) {
		wasmRangeOps = initWasmRangeOps()
	}
	return wasmRangeOps
}

export function resetWasmRangeOpsForTest(): void {
	wasmRangeOps = undefined
}
