import type { StyleId } from './ids.ts'
import type { CellStyle } from './style.ts'

export const DEFAULT_STYLE: CellStyle = Object.freeze({}) as CellStyle
const DEFAULT_STYLE_HASH = JSON.stringify(DEFAULT_STYLE)

export class StyleRegistry {
	private styles: CellStyle[] = [DEFAULT_STYLE]
	private hashes = new Map<string, StyleId>()
	private shared = false

	constructor() {
		this.hashes.set(DEFAULT_STYLE_HASH, 0 as StyleId)
	}

	register(style: CellStyle): StyleId {
		const hash = JSON.stringify(style)
		const existing = this.hashes.get(hash)
		if (existing !== undefined) return existing

		this.ensureWritable()
		const id = this.styles.length as StyleId
		this.styles.push(freezeDeep(structuredClone(style)))
		this.hashes.set(hash, id)
		return id
	}

	get(id: StyleId): CellStyle | undefined {
		return this.styles[id]
	}

	get size(): number {
		return this.styles.length
	}

	clone(): StyleRegistry {
		const clone = new StyleRegistry()
		clone.copyFrom(this)
		return clone
	}

	copyFrom(other: StyleRegistry): void {
		this.styles = other.styles
		this.hashes = other.hashes
		this.shared = true
		other.shared = true
	}

	private ensureWritable(): void {
		if (!this.shared) return
		this.styles = [...this.styles]
		this.hashes = new Map(this.hashes)
		this.shared = false
	}
}

function freezeDeep<T>(value: T): T {
	if (typeof value !== 'object' || value === null || Object.isFrozen(value)) return value
	Object.freeze(value)
	for (const child of Object.values(value as Record<string, unknown>)) {
		freezeDeep(child)
	}
	return value
}
