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

const A1_RE = /^([A-Za-z]+)(\d+)$/

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

	const bang = ref.indexOf('!')
	if (bang !== -1) {
		sheet = ref.substring(0, bang).replace(/^'|'$/g, '')
		body = ref.substring(bang + 1)
	}

	const parts = body.split(':')
	const startStr = parts[0]
	if (!startStr) throw new Error(`Invalid range reference: ${ref}`)
	const start = parseA1(startStr)
	const end = parts[1] ? parseA1(parts[1]) : start

	return sheet !== undefined ? { start, end, sheet } : { start, end }
}

export function toRangeString(ref: RangeRef): string {
	const range = `${toA1(ref.start)}:${toA1(ref.end)}`
	return ref.sheet ? `${ref.sheet}!${range}` : range
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
