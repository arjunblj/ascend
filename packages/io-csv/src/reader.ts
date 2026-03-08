import { createWorkbook, type Workbook } from '@ascend/core'
import {
	ascendError,
	booleanValue,
	type CellValue,
	type CsvDialect,
	EMPTY,
	err,
	numberValue,
	ok,
	type Result,
	stringValue,
} from '@ascend/schema'
import { resolveDialect } from './dialect.ts'

const BOM = '\uFEFF'

export function readCsv(content: string, dialect?: Partial<CsvDialect>): Result<Workbook> {
	const d = resolveDialect(dialect)
	const input = content.startsWith(BOM) ? content.slice(1) : content
	let rows: string[][]

	try {
		rows = parseFields(input, d)
	} catch (e) {
		return err(ascendError('IMPORT_ERROR', e instanceof Error ? e.message : 'CSV parse failed'))
	}

	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Sheet1')

	for (let r = 0; r < rows.length; r++) {
		const row = rows[r]
		if (!row) continue
		for (let c = 0; c < row.length; c++) {
			const raw = row[c] ?? ''
			const value = detectType(raw)
			if (value.kind !== 'empty') {
				sheet.cells.set(r, c, { value, formula: null, styleId: 0 as never })
			}
		}
	}

	return ok(workbook)
}

function parseFields(input: string, d: CsvDialect): string[][] {
	const rows: string[][] = []
	let row: string[] = []
	let i = 0
	const len = input.length
	const { delimiter, quote, escape: escapeChar } = d

	while (i <= len) {
		if (i === len) {
			if (row.length > 0 || rows.length > 0) {
				rows.push(row)
			}
			break
		}

		if (input[i] === quote) {
			let field = ''
			i++ // skip opening quote
			while (i < len) {
				if (input[i] === escapeChar && i + 1 < len && input[i + 1] === quote) {
					field += quote
					i += 2
				} else if (input[i] === quote) {
					i++ // skip closing quote
					break
				} else {
					field += input[i]
					i++
				}
			}
			row.push(field)
			if (i < len && input[i] === delimiter) {
				i++
			} else if (i < len && (input[i] === '\r' || input[i] === '\n')) {
				if (input[i] === '\r' && i + 1 < len && input[i + 1] === '\n') {
					i += 2
				} else {
					i++
				}
				rows.push(row)
				row = []
			}
		} else {
			let field = ''
			while (i < len && input[i] !== delimiter && input[i] !== '\r' && input[i] !== '\n') {
				field += input[i]
				i++
			}
			row.push(field)
			if (i < len && input[i] === delimiter) {
				i++
			} else {
				if (i < len && input[i] === '\r' && i + 1 < len && input[i + 1] === '\n') {
					i += 2
				} else if (i < len) {
					i++
				}
				rows.push(row)
				row = []
			}
		}
	}

	return rows
}

function detectType(raw: string): CellValue {
	if (raw === '') return EMPTY

	const lower = raw.toLowerCase()
	if (lower === 'true') return booleanValue(true)
	if (lower === 'false') return booleanValue(false)

	const num = Number(raw)
	if (raw.length > 0 && !Number.isNaN(num) && raw.trim() === raw) {
		return numberValue(num)
	}

	const date = tryParseDate(raw)
	if (date !== null) {
		return { kind: 'date', serial: date }
	}

	return stringValue(raw)
}

function tryParseDate(raw: string): number | null {
	const isoMatch = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$/.exec(raw)
	if (!isoMatch) return null

	const d = new Date(raw)
	if (Number.isNaN(d.getTime())) return null

	return dateToSerial(d)
}

function dateToSerial(d: Date): number {
	const epoch = new Date(1899, 11, 30)
	const diff = d.getTime() - epoch.getTime()
	return Math.round(diff / 86_400_000)
}
