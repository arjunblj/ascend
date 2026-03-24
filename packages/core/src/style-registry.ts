import { DEFAULT_STYLE_ID, type StyleId } from './ids.ts'
import type { BorderEdge, CellStyle, Color, GradientFill } from './style.ts'

export const DEFAULT_STYLE: CellStyle = Object.freeze({}) as CellStyle
const styleHashCache = new WeakMap<CellStyle, number>()
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
		if (fill.gradient) {
			h = fnv1aStr(h, fill.gradient.type ?? '')
			h = fnv1aNum(
				h,
				fill.gradient.degree !== undefined ? Math.round(fill.gradient.degree * 10000) : undefined,
			)
			h = fnv1aNum(
				h,
				fill.gradient.left !== undefined ? Math.round(fill.gradient.left * 10000) : undefined,
			)
			h = fnv1aNum(
				h,
				fill.gradient.right !== undefined ? Math.round(fill.gradient.right * 10000) : undefined,
			)
			h = fnv1aNum(
				h,
				fill.gradient.top !== undefined ? Math.round(fill.gradient.top * 10000) : undefined,
			)
			h = fnv1aNum(
				h,
				fill.gradient.bottom !== undefined ? Math.round(fill.gradient.bottom * 10000) : undefined,
			)
			for (const stop of fill.gradient.stops) {
				h = fnv1aNum(h, Math.round(stop.position * 10000))
				h = colorHash(h, stop.color)
			}
		}
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

function styleEquals(left: CellStyle, right: CellStyle): boolean {
	return (
		fontEquals(left.font, right.font) &&
		fillEquals(left.fill, right.fill) &&
		borderEquals(left.border, right.border) &&
		alignmentEquals(left.alignment, right.alignment) &&
		left.numberFormat === right.numberFormat &&
		protectionEquals(left.protection, right.protection)
	)
}

function fontEquals(
	left: CellStyle['font'] | undefined,
	right: CellStyle['font'] | undefined,
): boolean {
	return (
		left?.name === right?.name &&
		left?.size === right?.size &&
		left?.bold === right?.bold &&
		left?.italic === right?.italic &&
		left?.underline === right?.underline &&
		left?.strikethrough === right?.strikethrough &&
		colorEquals(left?.color, right?.color)
	)
}

function fillEquals(
	left: CellStyle['fill'] | undefined,
	right: CellStyle['fill'] | undefined,
): boolean {
	return (
		left?.pattern === right?.pattern &&
		colorEquals(left?.fgColor, right?.fgColor) &&
		colorEquals(left?.bgColor, right?.bgColor) &&
		gradientEquals(left?.gradient, right?.gradient)
	)
}

function borderEquals(
	left: CellStyle['border'] | undefined,
	right: CellStyle['border'] | undefined,
): boolean {
	return (
		borderEdgeEquals(left?.top, right?.top) &&
		borderEdgeEquals(left?.bottom, right?.bottom) &&
		borderEdgeEquals(left?.left, right?.left) &&
		borderEdgeEquals(left?.right, right?.right) &&
		borderEdgeEquals(left?.diagonal, right?.diagonal) &&
		left?.diagonalUp === right?.diagonalUp &&
		left?.diagonalDown === right?.diagonalDown
	)
}

function alignmentEquals(
	left: CellStyle['alignment'] | undefined,
	right: CellStyle['alignment'] | undefined,
): boolean {
	return (
		left?.horizontal === right?.horizontal &&
		left?.vertical === right?.vertical &&
		left?.wrapText === right?.wrapText &&
		left?.shrinkToFit === right?.shrinkToFit &&
		left?.textRotation === right?.textRotation &&
		left?.indent === right?.indent &&
		left?.readingOrder === right?.readingOrder
	)
}

function protectionEquals(
	left: CellStyle['protection'] | undefined,
	right: CellStyle['protection'] | undefined,
): boolean {
	return left?.locked === right?.locked && left?.hidden === right?.hidden
}

function gradientEquals(left: GradientFill | undefined, right: GradientFill | undefined): boolean {
	if (left === right) return true
	if (!left || !right) return false
	if (
		left.type !== right.type ||
		left.degree !== right.degree ||
		left.left !== right.left ||
		left.right !== right.right ||
		left.top !== right.top ||
		left.bottom !== right.bottom ||
		left.stops.length !== right.stops.length
	) {
		return false
	}
	for (let i = 0; i < left.stops.length; i++) {
		const leftStop = left.stops[i]
		const rightStop = right.stops[i]
		if (!leftStop || !rightStop) return false
		if (leftStop.position !== rightStop.position || !colorEquals(leftStop.color, rightStop.color)) {
			return false
		}
	}
	return true
}

function colorEquals(left: Color | undefined, right: Color | undefined): boolean {
	if (left === right) return true
	if (!left || !right || left.kind !== right.kind) return false
	switch (left.kind) {
		case 'theme':
			return right.kind === 'theme' && left.theme === right.theme && left.tint === right.tint
		case 'rgb':
			return right.kind === 'rgb' && left.rgb === right.rgb
		case 'indexed':
			return right.kind === 'indexed' && left.index === right.index
		default:
			return true
	}
}

function borderEdgeEquals(left: BorderEdge | undefined, right: BorderEdge | undefined): boolean {
	return left?.style === right?.style && colorEquals(left?.color, right?.color)
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

function borderEdgeHash(h: number, e: BorderEdge | undefined): number {
	if (!e) return h
	h = fnv1aStr(h, e.style ?? '')
	return colorHash(h, e.color)
}

function cloneColor(c: Color): Color {
	return { ...c }
}

function cloneBorderEdge(e: BorderEdge): BorderEdge {
	return e.color ? { ...e, color: cloneColor(e.color) } : { ...e }
}

export function cloneStyle(style: CellStyle): CellStyle {
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
		if (style.fill.gradient) {
			fill.gradient = {
				...style.fill.gradient,
				stops: style.fill.gradient.stops.map((stop) => ({
					...stop,
					color: cloneColor(stop.color),
				})),
			}
		}
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
		this.hashBuckets.set(DEFAULT_STYLE_HASH, [DEFAULT_STYLE_ID])
	}

	register(style: CellStyle): StyleId {
		let hash = styleHashCache.get(style)
		if (hash === undefined) {
			hash = styleHash(style)
			styleHashCache.set(style, hash)
		}
		const bucket = this.hashBuckets.get(hash)
		if (bucket) {
			for (const id of bucket) {
				const s = this.styles[id]
				if (s && styleEquals(s, style)) return id
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
