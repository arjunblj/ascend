import { parseRange, type RangeRef, type Workbook } from '@ascend/core'
import { ascendError, type CellValue, type CsvDialect, err, ok, type Result } from '@ascend/schema'
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

	for (let r = range.start.row; r <= range.end.row; r++) {
		const fields: string[] = []
		for (let c = range.start.col; c <= range.end.col; c++) {
			const cell = sheet.cells.get(r, c)
			const text = cell ? formatValue(cell.value) : ''
			fields.push(quoteField(text, d))
		}
		lines.push(fields.join(d.delimiter))
	}

	return ok(lines.join(d.lineEnding))
}

function formatValue(value: CellValue): string {
	switch (value.kind) {
		case 'empty':
			return ''
		case 'number':
			return String(value.value)
		case 'string':
			return value.value
		case 'boolean':
			return value.value ? 'TRUE' : 'FALSE'
		case 'error':
			return value.value
		case 'date':
			return serialToIso(value.serial)
		case 'richText':
			return value.runs.map((r) => r.text).join('')
	}
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
