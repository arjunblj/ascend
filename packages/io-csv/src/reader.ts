import { createWorkbook, DEFAULT_STYLE_ID, type Sheet, type Workbook } from '@ascend/core'
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
		if (!parseUnquotedFieldsFast(input, d, sheet)) {
			parseFields(input, d, (row, r) => {
				for (let c = 0; c < row.length; c++) {
					const raw = row[c] ?? ''
					const value = detectType(raw)
					if (value.kind !== 'empty') {
						sheet.cells.setResolved(r, c, value, null, DEFAULT_STYLE_ID)
					}
				}
			})
		}
	} catch (e) {
		return err(ascendError('IMPORT_ERROR', e instanceof Error ? e.message : 'CSV parse failed'))
	}

	return ok(workbook)
}

/**
 * Decode a UTF-8 byte stream into text and parse as CSV (full buffer after stream ends).
 * For chunked processing without building a `Workbook`, use application-level line splitting.
 */
export async function readCsvFromUtf8Stream(
	stream: ReadableStream<Uint8Array>,
	dialect?: Partial<CsvDialect>,
): Promise<Result<Workbook>> {
	const decoder = new TextDecoder('utf-8')
	let text = ''
	const reader = stream.getReader()
	for (;;) {
		const { done, value } = await reader.read()
		if (done) break
		if (value) text += decoder.decode(value, { stream: true })
	}
	text += decoder.decode()
	return readCsv(text, dialect)
}

type RowCallback = (row: string[], rowIndex: number) => void

function parseUnquotedFieldsFast(input: string, d: CsvDialect, sheet: Sheet): boolean {
	if (d.delimiter.length !== 1 || d.quote.length !== 1 || input.indexOf(d.quote) !== -1) {
		return false
	}

	const delimiterCode = d.delimiter.charCodeAt(0)
	let row = 0
	let col = 0
	let start = 0
	let i = 0
	const len = input.length

	while (i <= len) {
		if (i === len) {
			setUnquotedField(input, start, i, sheet, row, col)
			break
		}

		const code = input.charCodeAt(i)
		if (code === delimiterCode) {
			setUnquotedField(input, start, i, sheet, row, col)
			col++
			i++
			start = i
			continue
		}
		if (code === 13 || code === 10) {
			setUnquotedField(input, start, i, sheet, row, col)
			row++
			col = 0
			i += code === 13 && i + 1 < len && input.charCodeAt(i + 1) === 10 ? 2 : 1
			start = i
			continue
		}
		i++
	}

	return true
}

function setUnquotedField(
	input: string,
	start: number,
	end: number,
	sheet: Sheet,
	row: number,
	col: number,
): void {
	if (start === end) return

	if (
		!isTrimWhitespaceCode(input.charCodeAt(start)) &&
		!isTrimWhitespaceCode(input.charCodeAt(end - 1))
	) {
		const fastNumber = parseSimpleNumber(input, start, end)
		if (fastNumber !== null) {
			sheet.cells.setPlainNumber(row, col, fastNumber)
			return
		}
		const raw = input.slice(start, end)
		const num = Number(raw)
		if (!Number.isNaN(num)) {
			sheet.cells.setPlainNumber(row, col, num)
			return
		}
		if (equalsAsciiIgnoreCase(raw, 'true')) {
			sheet.cells.setResolved(row, col, booleanValue(true), null, DEFAULT_STYLE_ID)
			return
		}
		if (equalsAsciiIgnoreCase(raw, 'false')) {
			sheet.cells.setResolved(row, col, booleanValue(false), null, DEFAULT_STYLE_ID)
			return
		}
		const date = tryParseDate(raw)
		if (date !== null) {
			sheet.cells.setResolved(row, col, dateValue(date), null, DEFAULT_STYLE_ID)
			return
		}
		sheet.cells.setPlainString(row, col, raw)
		return
	}

	sheet.cells.setPlainString(row, col, input.slice(start, end))
}

function parseSimpleNumber(input: string, start: number, end: number): number | null {
	let i = start
	let sign = 1
	const first = input.charCodeAt(i)
	if (first === 45 || first === 43) {
		sign = first === 45 ? -1 : 1
		i++
		if (i === end) return null
	}

	let value = 0
	let hasDigit = false
	while (i < end) {
		const digit = input.charCodeAt(i) - 48
		if (digit < 0 || digit > 9) break
		value = value * 10 + digit
		hasDigit = true
		i++
	}

	if (i < end && input.charCodeAt(i) === 46) {
		i++
		let scale = 0.1
		while (i < end) {
			const digit = input.charCodeAt(i) - 48
			if (digit < 0 || digit > 9) break
			value += digit * scale
			scale *= 0.1
			hasDigit = true
			i++
		}
	}

	if (!hasDigit) return null

	if (i < end) {
		const exponentMarker = input.charCodeAt(i)
		if (exponentMarker !== 69 && exponentMarker !== 101) return null
		i++
		let exponentSign = 1
		const signCode = input.charCodeAt(i)
		if (signCode === 45 || signCode === 43) {
			exponentSign = signCode === 45 ? -1 : 1
			i++
		}
		let exponent = 0
		let hasExponentDigit = false
		while (i < end) {
			const digit = input.charCodeAt(i) - 48
			if (digit < 0 || digit > 9) return null
			exponent = exponent * 10 + digit
			hasExponentDigit = true
			i++
		}
		if (!hasExponentDigit) return null
		value *= 10 ** (exponentSign * exponent)
	}

	return i === end ? sign * value : null
}

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

	if (!hasEdgeTrimWhitespace(raw)) {
		const num = Number(raw)
		if (!Number.isNaN(num)) {
			return numberValue(num)
		}
	}

	if (equalsAsciiIgnoreCase(raw, 'true')) return booleanValue(true)
	if (equalsAsciiIgnoreCase(raw, 'false')) return booleanValue(false)

	const date = tryParseDate(raw)
	if (date !== null) {
		return dateValue(date)
	}

	return stringValue(raw)
}

function hasEdgeTrimWhitespace(raw: string): boolean {
	return (
		isTrimWhitespaceCode(raw.charCodeAt(0)) || isTrimWhitespaceCode(raw.charCodeAt(raw.length - 1))
	)
}

function isTrimWhitespaceCode(code: number): boolean {
	return (
		(code >= 0x09 && code <= 0x0d) ||
		code === 0x20 ||
		code === 0xa0 ||
		code === 0x1680 ||
		(code >= 0x2000 && code <= 0x200a) ||
		code === 0x2028 ||
		code === 0x2029 ||
		code === 0x202f ||
		code === 0x205f ||
		code === 0x3000 ||
		code === 0xfeff
	)
}

function equalsAsciiIgnoreCase(raw: string, literal: string): boolean {
	if (raw.length !== literal.length) return false
	for (let i = 0; i < literal.length; i++) {
		let code = raw.charCodeAt(i)
		if (code >= 65 && code <= 90) code += 32
		if (code !== literal.charCodeAt(i)) return false
	}
	return true
}

function tryParseDate(raw: string): number | null {
	const isoMatch = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$/.exec(raw)
	if (!isoMatch) return null

	const d = new Date(raw)
	if (Number.isNaN(d.getTime())) return null

	return dateToSerial(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate())
}
