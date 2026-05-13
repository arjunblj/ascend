export interface CellRef {
	readonly row: number
	readonly col: number
}

export interface RangeRef {
	readonly start: CellRef
	readonly end: CellRef
	readonly sheet?: string
}

const COLUMN_TO_INDEX_CACHE = new Map<string, number>()

export function columnToIndex(col: string): number {
	const upper = col.toUpperCase()
	const cached = COLUMN_TO_INDEX_CACHE.get(upper)
	if (cached !== undefined) return cached
	let result = 0
	for (let i = 0; i < upper.length; i++) {
		result = result * 26 + (upper.charCodeAt(i) - 64)
	}
	const index = result - 1
	if (index >= 0 && index < 702) COLUMN_TO_INDEX_CACHE.set(upper, index)
	return index
}

const COLUMN_CACHE: string[] = []

function computeColumnLabel(index: number): string {
	let result = ''
	let n = index + 1
	while (n > 0) {
		const rem = (n - 1) % 26
		result = String.fromCharCode(65 + rem) + result
		n = Math.floor((n - 1) / 26)
	}
	return result
}

for (let i = 0; i < 702; i++) {
	const label = computeColumnLabel(i)
	COLUMN_CACHE.push(label)
	COLUMN_TO_INDEX_CACHE.set(label, i)
}

export function indexToColumn(index: number): string {
	return COLUMN_CACHE[index] ?? computeColumnLabel(index)
}

const MAX_EXCEL_ROW_INDEX = 1_048_575
const MAX_EXCEL_COL_INDEX = 16_383

const A1_RE = /^\$?([A-Za-z]+)\$?(\d+)$/
const WHOLE_COLUMN_RE = /^\$?([A-Za-z]+)$/
const WHOLE_ROW_RE = /^\$?(\d+)$/

export function parseA1Safe(ref: string | undefined): CellRef | null {
	if (!ref) return null
	const m = A1_RE.exec(ref)
	if (!m?.[1] || !m[2]) return null
	return {
		row: Number.parseInt(m[2], 10) - 1,
		col: columnToIndex(m[1].toUpperCase()),
	}
}

export function parseA1(ref: string): CellRef {
	const parsed = parseA1Safe(ref)
	if (!parsed) throw new Error(`Invalid A1 reference: ${ref}`)
	return parsed
}

export function toA1(ref: CellRef): string {
	return indexToColumn(ref.col) + (ref.row + 1)
}

export function parseRange(ref: string): RangeRef {
	let sheet: string | undefined
	let body = ref

	const bang = findSheetSeparator(ref)
	if (bang !== -1) {
		sheet = ref.substring(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'")
		body = ref.substring(bang + 1)
	}

	const parts = body.split(':')
	if (parts.length > 2 || parts.some((part) => part === '')) {
		throw new Error(`Invalid range reference: ${ref}`)
	}
	const startStr = parts[0]
	if (!startStr) throw new Error(`Invalid range reference: ${ref}`)
	const endStr = parts[1]
	let start: CellRef
	let end: CellRef
	if (endStr && WHOLE_ROW_RE.test(startStr) && WHOLE_ROW_RE.test(endStr)) {
		start = { row: Number.parseInt(startStr.replace('$', ''), 10) - 1, col: 0 }
		end = { row: Number.parseInt(endStr.replace('$', ''), 10) - 1, col: MAX_EXCEL_COL_INDEX }
	} else if (endStr && WHOLE_COLUMN_RE.test(startStr) && WHOLE_COLUMN_RE.test(endStr)) {
		start = { row: 0, col: columnToIndex(startStr.replace('$', '').toUpperCase()) }
		end = { row: MAX_EXCEL_ROW_INDEX, col: columnToIndex(endStr.replace('$', '').toUpperCase()) }
	} else {
		start = parseA1(startStr)
		end = endStr ? parseA1(endStr) : start
	}

	return sheet !== undefined ? { start, end, sheet } : { start, end }
}

export function toRangeString(ref: RangeRef): string {
	const range = `${toA1(ref.start)}:${toA1(ref.end)}`
	return ref.sheet ? `${formatSheetName(ref.sheet)}!${range}` : range
}

export function normalizeRange(range: RangeRef): RangeRef {
	const start = {
		row: Math.min(range.start.row, range.end.row),
		col: Math.min(range.start.col, range.end.col),
	}
	const end = {
		row: Math.max(range.start.row, range.end.row),
		col: Math.max(range.start.col, range.end.col),
	}
	return range.sheet !== undefined ? { start, end, sheet: range.sheet } : { start, end }
}

export function rangeIntersects(left: RangeRef, right: RangeRef): boolean {
	if (left.sheet !== undefined && right.sheet !== undefined && left.sheet !== right.sheet) {
		return false
	}
	const a = normalizeRange(left)
	const b = normalizeRange(right)
	return (
		a.start.row <= b.end.row &&
		a.end.row >= b.start.row &&
		a.start.col <= b.end.col &&
		a.end.col >= b.start.col
	)
}

export function rangeIntersection(left: RangeRef, right: RangeRef): RangeRef | null {
	if (!rangeIntersects(left, right)) return null
	const a = normalizeRange(left)
	const b = normalizeRange(right)
	const sheet = a.sheet ?? b.sheet
	const range = {
		start: {
			row: Math.max(a.start.row, b.start.row),
			col: Math.max(a.start.col, b.start.col),
		},
		end: {
			row: Math.min(a.end.row, b.end.row),
			col: Math.min(a.end.col, b.end.col),
		},
	}
	return sheet !== undefined ? { ...range, sheet } : range
}

export function parseSqref(sqref: string): RangeRef[] {
	return splitSqref(sqref).map(parseRange)
}

export function sqrefIntersects(sqref: string, range: RangeRef): boolean {
	return parseSqref(sqref).some((candidate) => rangeIntersects(candidate, range))
}

function findSheetSeparator(ref: string): number {
	let quoted = false
	for (let index = 0; index < ref.length; index++) {
		const ch = ref[index]
		if (ch === "'") {
			if (quoted && ref[index + 1] === "'") {
				index++
			} else {
				quoted = !quoted
			}
		} else if (ch === '!' && !quoted) {
			return index
		}
	}
	return -1
}

function formatSheetName(sheet: string): string {
	if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet)) return sheet
	return `'${sheet.replace(/'/g, "''")}'`
}

function splitSqref(sqref: string): string[] {
	const refs: string[] = []
	let current = ''
	let quoted = false
	for (let index = 0; index < sqref.length; index++) {
		const ch = sqref[index] ?? ''
		if (ch === "'") {
			current += ch
			if (quoted && sqref[index + 1] === "'") {
				current += "'"
				index++
			} else {
				quoted = !quoted
			}
		} else if (/\s/.test(ch) && !quoted) {
			if (current) {
				refs.push(current)
				current = ''
			}
		} else {
			current += ch
		}
	}
	if (current) refs.push(current)
	return refs
}

export function expandRange(range: RangeRef): CellRef[] {
	const cells: CellRef[] = []
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			cells.push({ row, col })
		}
	}
	return cells
}

export function forEachCellInRange(range: RangeRef, fn: (row: number, col: number) => void): void {
	for (let row = range.start.row; row <= range.end.row; row++) {
		for (let col = range.start.col; col <= range.end.col; col++) {
			fn(row, col)
		}
	}
}
