import type { StyleId } from './ids.ts'
import type { CellStyle } from './style.ts'

export const DEFAULT_STYLE: CellStyle = {}

export class StyleRegistry {
	private readonly styles: CellStyle[] = [DEFAULT_STYLE]
	private readonly hashes = new Map<string, StyleId>()

	constructor() {
		this.hashes.set(JSON.stringify(DEFAULT_STYLE), 0 as StyleId)
	}

	register(style: CellStyle): StyleId {
		const hash = JSON.stringify(style)
		const existing = this.hashes.get(hash)
		if (existing !== undefined) return existing

		const id = this.styles.length as StyleId
		this.styles.push(style)
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
		clone.styles.splice(
			0,
			clone.styles.length,
			...this.styles.map((style) => structuredClone(style)),
		)
		clone.hashes.clear()
		for (const [hash, id] of this.hashes) {
			clone.hashes.set(hash, id)
		}
		return clone
	}
}
