import type { StyleId } from './ids.ts'
import type { BorderEdge, CellStyle, Color } from './style.ts'

export const DEFAULT_STYLE: CellStyle = Object.freeze({}) as CellStyle
const DEFAULT_STYLE_HASH = styleHash(DEFAULT_STYLE)

function styleHash(style: CellStyle): string {
	const font = style.font
	const fill = style.fill
	const border = style.border
	const alignment = style.alignment
	const protection = style.protection
	const parts: string[] = [
		font?.name ?? '',
		String(font?.size ?? ''),
		String(font?.bold ?? ''),
		String(font?.italic ?? ''),
		String(font?.underline ?? ''),
		String(font?.strikethrough ?? ''),
		colorKey(font?.color),
		fill?.pattern ?? '',
		colorKey(fill?.fgColor),
		colorKey(fill?.bgColor),
		borderEdgeKey(border?.top),
		borderEdgeKey(border?.bottom),
		borderEdgeKey(border?.left),
		borderEdgeKey(border?.right),
		borderEdgeKey(border?.diagonal),
		String(border?.diagonalUp ?? ''),
		String(border?.diagonalDown ?? ''),
		alignment?.horizontal ?? '',
		alignment?.vertical ?? '',
		String(alignment?.wrapText ?? ''),
		String(alignment?.shrinkToFit ?? ''),
		String(alignment?.textRotation ?? ''),
		String(alignment?.indent ?? ''),
		String(alignment?.readingOrder ?? ''),
		style.numberFormat ?? '',
		String(protection?.locked ?? ''),
		String(protection?.hidden ?? ''),
	]
	return parts.join('|')
}

function colorKey(c: Color | undefined): string {
	if (!c) return ''
	if (c.kind === 'theme') return `theme:${c.theme}:${c.tint ?? ''}`
	if (c.kind === 'rgb') return `rgb:${c.rgb}`
	if (c.kind === 'indexed') return `idx:${c.index}`
	return 'auto'
}

function borderEdgeKey(e: BorderEdge | undefined): string {
	if (!e) return ''
	return `${e.style ?? ''}:${colorKey(e.color)}`
}

export class StyleRegistry {
	private styles: CellStyle[] = [DEFAULT_STYLE]
	private hashes = new Map<string, StyleId>()
	private shared = false

	constructor() {
		this.hashes.set(DEFAULT_STYLE_HASH, 0 as StyleId)
	}

	register(style: CellStyle): StyleId {
		const hash = styleHash(style)
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
