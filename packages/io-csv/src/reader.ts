import { createWorkbook, DEFAULT_STYLE_ID, type Workbook } from '@ascend/core'
import { dateToSerial } from '@ascend/formulas'
import {
	ascendError,
	booleanValue,
	type CellValue,
	type CsvDialect,
	dateValue,
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

	const workbook = createWorkbook()
	const sheet = workbook.addSheet('Sheet1')

	try {
		parseFields(input, d, (row, r) => {
			for (let c = 0; c < row.length; c++) {
				const raw = row[c] ?? ''
				const value = detectType(raw)
				if (value.kind !== 'empty') {
					sheet.cells.setResolved(r, c, value, null, DEFAULT_STYLE_ID)
				}
			}
		})
	} catch (e) {
		return err(ascendError('IMPORT_ERROR', e instanceof Error ? e.message : 'CSV parse failed'))
	}

	return ok(workbook)
}

type RowCallback = (row: string[], rowIndex: number) => void

function parseFields(input: string, d: CsvDialect, onRow: RowCallback): void {
	let row: string[] = []
	let rowIndex = 0
	let i = 0
	const len = input.length
	const { delimiter, quote, escape: escapeChar } = d

	const emitRow = () => {
		onRow(row, rowIndex++)
		row = []
	}

	while (i <= len) {
		if (i === len) {
			if (row.length > 0 || rowIndex > 0) {
				emitRow()
			}
			break
		}

		if (input[i] === quote) {
			i++
			const start = i
			let hasEscape = false
			while (i < len) {
				if (input[i] === escapeChar && i + 1 < len && input[i + 1] === quote) {
					hasEscape = true
					i += 2
				} else if (input[i] === quote) {
					break
				} else {
					i++
				}
			}
			const raw = input.slice(start, i)
			row.push(hasEscape ? raw.replaceAll(escapeChar + quote, quote) : raw)
			if (i < len) i++
			if (i < len && input[i] === delimiter) {
				i++
			} else if (i < len && (input[i] === '\r' || input[i] === '\n')) {
				if (input[i] === '\r' && i + 1 < len && input[i + 1] === '\n') {
					i += 2
				} else {
					i++
				}
				emitRow()
			}
		} else {
			const start = i
			while (i < len && input[i] !== delimiter && input[i] !== '\r' && input[i] !== '\n') {
				i++
			}
			row.push(input.slice(start, i))
			if (i < len && input[i] === delimiter) {
				i++
			} else {
				if (i < len && input[i] === '\r' && i + 1 < len && input[i + 1] === '\n') {
					i += 2
				} else if (i < len) {
					i++
				}
				emitRow()
			}
		}
	}
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
		return dateValue(date)
	}

	return stringValue(raw)
}

function tryParseDate(raw: string): number | null {
	const isoMatch = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$/.exec(raw)
	if (!isoMatch) return null

	const d = new Date(raw)
	if (Number.isNaN(d.getTime())) return null

	return dateToSerial(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate())
}
