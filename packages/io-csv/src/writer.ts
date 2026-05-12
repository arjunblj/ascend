import { parseRange, type RangeRef, type Workbook } from '@ascend/core'
import {
	ascendError,
	type CellValue,
	type CsvDialect,
	err,
	ok,
	type Result,
	topLeftScalar,
} from '@ascend/schema'
import { resolveDialect } from './dialect.ts'

export interface WriteCsvOptions {
	readonly sheet?: string
	readonly range?: string
	readonly dialect?: Partial<CsvDialect>
}

export function writeCsv(workbook: Workbook, opts?: WriteCsvOptions): Result<string> {
	const d = resolveDialect(opts?.dialect)

	const sheet = opts?.sheet ? workbook.getSheet(opts.sheet) : workbook.sheets[0]

	if (!sheet) {
		return err(
			ascendError('SHEET_NOT_FOUND', `Sheet not found: ${opts?.sheet ?? '(empty workbook)'}`),
		)
	}

	let range: RangeRef
	if (opts?.range) {
		try {
			range = parseRange(opts.range)
		} catch {
			return err(ascendError('INVALID_RANGE', `Invalid range: ${opts.range}`))
		}
	} else {
		const used = sheet.cells.usedRange()
		if (!used) return ok('')
		range = used
	}

	const lines: string[] = []
	const plainScalarFastPath = d.delimiter === ',' && d.quote === '"' && d.escape === '"'

	for (let r = range.start.row; r <= range.end.row; r++) {
		const fields: string[] = []
		for (let c = range.start.col; c <= range.end.col; c++) {
			fields.push(formatField(sheet.cells.readValue(r, c), d, plainScalarFastPath))
		}
		lines.push(fields.join(d.delimiter))
	}

	return ok(lines.join(d.lineEnding))
}

function formatField(value: CellValue, d: CsvDialect, plainScalarFastPath: boolean): string {
	value = topLeftScalar(value)
	switch (value.kind) {
		case 'empty':
			return ''
		case 'number':
			return formatScalarField(String(value.value), d, plainScalarFastPath)
		case 'string':
			return quoteField(value.value, d)
		case 'boolean':
			return formatScalarField(value.value ? 'TRUE' : 'FALSE', d, plainScalarFastPath)
		case 'error':
			return formatScalarField(value.value, d, plainScalarFastPath)
		case 'date':
			return formatScalarField(serialToIso(value.serial), d, plainScalarFastPath)
		case 'richText':
			return quoteField(value.runs.map((r) => r.text).join(''), d)
	}
}

function formatScalarField(text: string, d: CsvDialect, plainScalarFastPath: boolean): string {
	return plainScalarFastPath ? text : quoteField(text, d)
}

function quoteField(text: string, d: CsvDialect): string {
	if (
		text.includes(d.delimiter) ||
		text.includes(d.quote) ||
		text.includes('\n') ||
		text.includes('\r')
	) {
		const escaped = text.replaceAll(d.quote, d.escape + d.quote)
		return d.quote + escaped + d.quote
	}
	return text
}

function serialToIso(serial: number): string {
	const epoch = new Date(1899, 11, 30)
	const ms = epoch.getTime() + serial * 86_400_000
	const d = new Date(ms)
	const y = d.getFullYear()
	const m = String(d.getMonth() + 1).padStart(2, '0')
	const day = String(d.getDate()).padStart(2, '0')
	return `${y}-${m}-${day}`
}
