import type { StyleId } from './ids.ts'
import type { BorderEdge, CellStyle, Color } from './style.ts'

export const DEFAULT_STYLE: CellStyle = Object.freeze({}) as CellStyle
const DEFAULT_STYLE_HASH = styleHash(DEFAULT_STYLE)

function fnv1aStr(h: number, s: string): number {
	for (let i = 0; i < s.length; i++) {
		h ^= s.charCodeAt(i)
		h = Math.imul(h, 0x01000193)
	}
	return h
}

function fnv1aBool(h: number, v: boolean | undefined): number {
	h ^= v ? 1 : 0
	return Math.imul(h, 0x01000193)
}

function fnv1aNum(h: number, v: number | undefined): number {
	if (v === undefined) return h
	h ^= (v | 0) & 0xff
	h = Math.imul(h, 0x01000193)
	h ^= ((v | 0) >> 8) & 0xff
	return Math.imul(h, 0x01000193)
}

function styleHash(style: CellStyle): number {
	let h = 0x811c9dc5
	const font = style.font
	const fill = style.fill
	const border = style.border
	const alignment = style.alignment
	const protection = style.protection
	if (font) {
		h = fnv1aStr(h, font.name ?? '')
		h = fnv1aNum(h, font.size)
		h = fnv1aBool(h, font.bold)
		h = fnv1aBool(h, font.italic)
		h = fnv1aStr(h, String(font.underline ?? ''))
		h = fnv1aBool(h, font.strikethrough)
		h = colorHash(h, font.color)
	}
	if (fill) {
		h = fnv1aStr(h, fill.pattern ?? '')
		h = colorHash(h, fill.fgColor)
		h = colorHash(h, fill.bgColor)
	}
	if (border) {
		h = borderEdgeHash(h, border.top)
		h = borderEdgeHash(h, border.bottom)
		h = borderEdgeHash(h, border.left)
		h = borderEdgeHash(h, border.right)
		h = borderEdgeHash(h, border.diagonal)
		h = fnv1aBool(h, border.diagonalUp)
		h = fnv1aBool(h, border.diagonalDown)
	}
	if (alignment) {
		h = fnv1aStr(h, alignment.horizontal ?? '')
		h = fnv1aStr(h, alignment.vertical ?? '')
		h = fnv1aBool(h, alignment.wrapText)
		h = fnv1aBool(h, alignment.shrinkToFit)
		h = fnv1aNum(h, alignment.textRotation)
		h = fnv1aNum(h, alignment.indent)
		h = fnv1aNum(h, alignment.readingOrder)
	}
	h = fnv1aStr(h, style.numberFormat ?? '')
	if (protection) {
		h = fnv1aBool(h, protection.locked)
		h = fnv1aBool(h, protection.hidden)
	}
	return h >>> 0
}

function styleFingerprint(style: CellStyle): string {
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

function colorHash(h: number, c: Color | undefined): number {
	if (!c) return h
	h = fnv1aStr(h, c.kind)
	if (c.kind === 'theme') {
		h = fnv1aNum(h, c.theme)
		if (c.tint !== undefined) h = fnv1aNum(h, Math.round(c.tint * 10000))
	} else if (c.kind === 'rgb') {
		h = fnv1aStr(h, c.rgb)
	} else if (c.kind === 'indexed') {
		h = fnv1aNum(h, c.index)
	}
	return h
}

function colorKey(c: Color | undefined): string {
	if (!c) return ''
	if (c.kind === 'theme') return `theme:${c.theme}:${c.tint ?? ''}`
	if (c.kind === 'rgb') return `rgb:${c.rgb}`
	if (c.kind === 'indexed') return `idx:${c.index}`
	return 'auto'
}

function borderEdgeHash(h: number, e: BorderEdge | undefined): number {
	if (!e) return h
	h = fnv1aStr(h, e.style ?? '')
	return colorHash(h, e.color)
}

function borderEdgeKey(e: BorderEdge | undefined): string {
	if (!e) return ''
	return `${e.style ?? ''}:${colorKey(e.color)}`
}

function cloneColor(c: Color): Color {
	return { ...c }
}

function cloneBorderEdge(e: BorderEdge): BorderEdge {
	return e.color ? { ...e, color: cloneColor(e.color) } : { ...e }
}

function cloneStyle(style: CellStyle): CellStyle {
	const result: Record<string, unknown> = {}
	if (style.font) {
		result.font = style.font.color
			? { ...style.font, color: cloneColor(style.font.color) }
			: { ...style.font }
	}
	if (style.fill) {
		const fill: Record<string, unknown> = { ...style.fill }
		if (style.fill.fgColor) fill.fgColor = cloneColor(style.fill.fgColor)
		if (style.fill.bgColor) fill.bgColor = cloneColor(style.fill.bgColor)
		result.fill = fill
	}
	if (style.border) {
		const border: Record<string, unknown> = {}
		if (style.border.top) border.top = cloneBorderEdge(style.border.top)
		if (style.border.bottom) border.bottom = cloneBorderEdge(style.border.bottom)
		if (style.border.left) border.left = cloneBorderEdge(style.border.left)
		if (style.border.right) border.right = cloneBorderEdge(style.border.right)
		if (style.border.diagonal) border.diagonal = cloneBorderEdge(style.border.diagonal)
		if (style.border.diagonalUp !== undefined) border.diagonalUp = style.border.diagonalUp
		if (style.border.diagonalDown !== undefined) border.diagonalDown = style.border.diagonalDown
		result.border = border
	}
	if (style.alignment) result.alignment = { ...style.alignment }
	if (style.numberFormat !== undefined) result.numberFormat = style.numberFormat
	if (style.protection) result.protection = { ...style.protection }
	return result as CellStyle
}

export class StyleRegistry {
	private styles: CellStyle[] = [DEFAULT_STYLE]
	private hashBuckets = new Map<number, StyleId[]>()
	private shared = false

	constructor() {
		this.hashBuckets.set(DEFAULT_STYLE_HASH, [0 as StyleId])
	}

	register(style: CellStyle): StyleId {
		const hash = styleHash(style)
		const bucket = this.hashBuckets.get(hash)
		if (bucket) {
			const fp = styleFingerprint(style)
			for (const id of bucket) {
				const s = this.styles[id]
				if (s && styleFingerprint(s) === fp) return id
			}
		}

		this.ensureWritable()
		const id = this.styles.length as StyleId
		this.styles.push(freezeDeep(cloneStyle(style)))
		let b = this.hashBuckets.get(hash)
		if (!b) {
			b = []
			this.hashBuckets.set(hash, b)
		}
		b.push(id)
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
		this.hashBuckets = other.hashBuckets
		this.shared = true
		other.shared = true
	}

	private ensureWritable(): void {
		if (!this.shared) return
		this.styles = [...this.styles]
		const prev = this.hashBuckets
		this.hashBuckets = new Map()
		for (const [hash, bucket] of prev) {
			this.hashBuckets.set(hash, [...bucket])
		}
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
