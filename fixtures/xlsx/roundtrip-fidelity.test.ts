import { describe, expect, test } from 'bun:test'
import type { StyleId } from '../../packages/core/src/index.ts'
import { createTableId, Workbook } from '../../packages/core/src/index.ts'
import { readXlsx } from '../../packages/io-xlsx/src/reader/index.ts'
import { writeXlsx } from '../../packages/io-xlsx/src/writer/index.ts'
import {
	booleanValue,
	errorValue,
	numberValue,
	stringValue,
} from '../../packages/schema/src/index.ts'

const S0 = 0 as StyleId

function roundTrip(wb: Workbook) {
	const written = writeXlsx(wb)
	if (!written.ok) throw new Error(`write failed: ${written.error.message}`)
	const read = readXlsx(written.value)
	if (!read.ok) throw new Error(`read failed: ${read.error.message}`)
	return { bytes: written.value, wb: read.value.workbook }
}

let original: Workbook
let result: Workbook

function sheetAt(index: number) {
	const sheet = result.sheets[index]
	if (!sheet) throw new Error(`Missing sheet at index ${index}`)
	return sheet
}

function must<T>(value: T | undefined | null, message: string): T {
	if (value === undefined || value === null) throw new Error(message)
	return value
}

describe('roundtrip fidelity', () => {
	test('setup: build and roundtrip a complex workbook', () => {
		original = new Workbook()

		// --- Sheet 1: Data ---
		const data = original.addSheet('Data')

		// Cell types
		data.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		data.cells.set(0, 1, { value: stringValue('Score'), formula: null, styleId: S0 })
		data.cells.set(0, 2, { value: stringValue('Pass'), formula: null, styleId: S0 })
		data.cells.set(0, 3, { value: stringValue('Date'), formula: null, styleId: S0 })
		data.cells.set(0, 4, { value: stringValue('Total'), formula: null, styleId: S0 })

		data.cells.set(1, 0, { value: stringValue('Alice'), formula: null, styleId: S0 })
		data.cells.set(1, 1, { value: numberValue(85), formula: null, styleId: S0 })
		data.cells.set(1, 2, { value: booleanValue(true), formula: null, styleId: S0 })
		data.cells.set(1, 3, { value: numberValue(44927), formula: null, styleId: S0 }) // 2023-01-01 as serial
		data.cells.set(1, 4, { value: numberValue(85), formula: 'SUM(B2)', styleId: S0 })

		data.cells.set(2, 0, { value: stringValue('Bob'), formula: null, styleId: S0 })
		data.cells.set(2, 1, { value: numberValue(92), formula: null, styleId: S0 })
		data.cells.set(2, 2, { value: booleanValue(true), formula: null, styleId: S0 })
		data.cells.set(2, 3, { value: numberValue(44958), formula: null, styleId: S0 }) // 2023-02-01
		data.cells.set(2, 4, { value: numberValue(92), formula: 'SUM(B3)', styleId: S0 })

		data.cells.set(3, 0, { value: stringValue('Carol'), formula: null, styleId: S0 })
		data.cells.set(3, 1, { value: numberValue(78), formula: null, styleId: S0 })
		data.cells.set(3, 2, { value: booleanValue(false), formula: null, styleId: S0 })
		data.cells.set(3, 3, { value: numberValue(44986), formula: null, styleId: S0 }) // 2023-03-01
		data.cells.set(3, 4, { value: numberValue(78), formula: 'SUM(B4)', styleId: S0 })

		// Error cell
		data.cells.set(4, 0, { value: errorValue('#N/A'), formula: null, styleId: S0 })

		// IF formula
		data.cells.set(5, 0, {
			value: stringValue('Pass'),
			formula: 'IF(B2>80,"Pass","Fail")',
			styleId: S0,
		})

		// Shared formulas (column E: =SUM(B<row>))
		data.cells.set(1, 4, {
			value: numberValue(85),
			formula: 'SUM(B2)',
			styleId: S0,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: '0',
				isMaster: true,
				masterRef: 'E2',
				sharedRange: 'E2:E4',
			},
		})
		data.cells.set(2, 4, {
			value: numberValue(92),
			formula: null,
			styleId: S0,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'E2' },
		})
		data.cells.set(3, 4, {
			value: numberValue(78),
			formula: null,
			styleId: S0,
			formulaInfo: { kind: 'shared', sharedIndex: '0', isMaster: false, masterRef: 'E2' },
		})

		// --- Sheet 2: Summary ---
		const summary = original.addSheet('Summary')
		// Cross-sheet ref
		summary.cells.set(0, 0, {
			value: numberValue(255),
			formula: 'SUM(Data!B2:B4)',
			styleId: S0,
		})
		// VLOOKUP
		summary.cells.set(1, 0, {
			value: numberValue(92),
			formula: 'VLOOKUP("Bob",Data!A2:B4,2,FALSE)',
			styleId: S0,
		})
		// Simple math
		summary.cells.set(2, 0, {
			value: numberValue(5),
			formula: '2+3',
			styleId: S0,
		})

		// --- Styles ---
		const boldId = original.styles.register({ font: { bold: true } })
		const italicId = original.styles.register({ font: { italic: true } })
		const coloredFontId = original.styles.register({
			font: { color: { kind: 'rgb', rgb: 'FFFF0000' } },
		})
		const fillId = original.styles.register({
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FF00FF00' } },
		})
		const borderId = original.styles.register({
			border: {
				top: { style: 'thin', color: { kind: 'rgb', rgb: 'FF000000' } },
				bottom: { style: 'thin', color: { kind: 'rgb', rgb: 'FF000000' } },
			},
		})
		const numFmtId = original.styles.register({ numberFormat: '0.00%' })
		const alignId = original.styles.register({
			alignment: { horizontal: 'center', vertical: 'top', wrapText: true },
		})

		// Apply styles to header row of Data sheet
		data.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: boldId })
		data.cells.set(0, 1, { value: stringValue('Score'), formula: null, styleId: italicId })

		// Sheet 3 for style-heavy cells
		const styles = original.addSheet('Styles')
		styles.cells.set(0, 0, { value: stringValue('Bold'), formula: null, styleId: boldId })
		styles.cells.set(0, 1, { value: stringValue('Italic'), formula: null, styleId: italicId })
		styles.cells.set(0, 2, { value: stringValue('Red'), formula: null, styleId: coloredFontId })
		styles.cells.set(0, 3, { value: numberValue(0.75), formula: null, styleId: numFmtId })
		styles.cells.set(1, 0, { value: stringValue('Green BG'), formula: null, styleId: fillId })
		styles.cells.set(1, 1, { value: stringValue('Bordered'), formula: null, styleId: borderId })
		styles.cells.set(1, 2, { value: stringValue('Centered'), formula: null, styleId: alignId })

		// --- Merged cells ---
		data.merges.push({ start: { row: 6, col: 0 }, end: { row: 6, col: 2 } })
		data.merges.push({ start: { row: 7, col: 0 }, end: { row: 8, col: 0 } })

		// --- Frozen panes ---
		data.frozenRows = 1
		data.frozenCols = 1

		// --- Column widths and row heights ---
		data.colWidths.set(0, 20)
		data.colWidths.set(1, 15)
		data.rowHeights.set(0, 30)
		data.rowHeights.set(5, 25)

		// --- Hidden rows and columns ---
		data.rowDefs.set(4, { hidden: true })
		data.colDefs.push({ min: 0, max: 0, width: 20, customWidth: true })
		data.colDefs.push({ min: 1, max: 1, width: 15, customWidth: true })
		data.colDefs.push({ min: 3, max: 3, hidden: true })

		// --- Conditional formatting ---
		original.differentialStyles.push({
			font: { bold: true },
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
		})
		original.differentialStyles.push({
			font: { italic: true },
			fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFFFC7CE' } },
		})
		data.conditionalFormats.push({
			sqref: 'B2:B4',
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					dxfId: 0,
					priority: 1,
					formulas: ['80'],
					style: original.differentialStyles[0],
				},
			],
		})
		data.conditionalFormats.push({
			sqref: 'B2:B4',
			rules: [
				{
					type: 'cellIs',
					operator: 'lessThan',
					dxfId: 1,
					priority: 2,
					formulas: ['80'],
					style: original.differentialStyles[1],
				},
			],
		})

		// --- Data validation ---
		data.dataValidations.push({
			sqref: 'C2:C4',
			type: 'list',
			formula1: '"Yes,No,Maybe"',
			allowBlank: true,
			showInputMessage: true,
		})
		data.dataValidations.push({
			sqref: 'B2:B4',
			type: 'whole',
			operator: 'between',
			formula1: '0',
			formula2: '100',
			allowBlank: true,
			showErrorMessage: true,
			errorTitle: 'Invalid Score',
			error: 'Score must be between 0 and 100',
		})

		// --- Comments ---
		data.comments.set('A1', { text: 'Student name', author: 'Teacher' })
		data.comments.set('B1', { text: 'Exam score out of 100' })

		// --- Hyperlinks ---
		data.hyperlinks.set('A1', {
			target: 'https://example.com/students',
			display: 'Name',
			tooltip: 'View students',
		})
		summary.hyperlinks.set('A1', {
			target: 'https://example.com/summary',
		})

		// --- Defined names ---
		original.definedNames.set('Scores', 'Data!$B$2:$B$4')
		original.definedNames.set('StudentNames', 'Data!$A$2:$A$4')

		// --- Auto-filter ---
		data.autoFilter = {
			ref: 'A1:E4',
			columns: [],
		}

		// --- Tables ---
		const tableSheet = original.addSheet('TableSheet')
		tableSheet.cells.set(0, 0, { value: stringValue('Product'), formula: null, styleId: S0 })
		tableSheet.cells.set(0, 1, { value: stringValue('Price'), formula: null, styleId: S0 })
		tableSheet.cells.set(1, 0, { value: stringValue('Widget'), formula: null, styleId: S0 })
		tableSheet.cells.set(1, 1, { value: numberValue(9.99), formula: null, styleId: S0 })
		tableSheet.cells.set(2, 0, { value: stringValue('Gadget'), formula: null, styleId: S0 })
		tableSheet.cells.set(2, 1, { value: numberValue(19.99), formula: null, styleId: S0 })
		tableSheet.tables.push({
			id: createTableId(),
			name: 'Products',
			sheetId: tableSheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Product' }, { name: 'Price' }],
			hasHeaders: true,
			hasTotals: false,
		})

		// --- Sheet protection ---
		summary.protection = {
			sheet: true,
			objects: true,
			scenarios: true,
			sort: false,
			autoFilter: false,
		}

		// Roundtrip
		const rt = roundTrip(original)
		result = rt.wb
	})

	test('multiple sheets with correct names', () => {
		expect(result.sheets).toHaveLength(4)
		expect(result.sheets[0]?.name).toBe('Data')
		expect(result.sheets[1]?.name).toBe('Summary')
		expect(result.sheets[2]?.name).toBe('Styles')
		expect(result.sheets[3]?.name).toBe('TableSheet')
	})

	test('string values', () => {
		const s = sheetAt(0)
		expect(s.cells.get(1, 0)?.value).toEqual({ kind: 'string', value: 'Alice' })
		expect(s.cells.get(2, 0)?.value).toEqual({ kind: 'string', value: 'Bob' })
		expect(s.cells.get(3, 0)?.value).toEqual({ kind: 'string', value: 'Carol' })
	})

	test('number values', () => {
		const s = sheetAt(0)
		expect(s.cells.get(1, 1)?.value).toEqual({ kind: 'number', value: 85 })
		expect(s.cells.get(2, 1)?.value).toEqual({ kind: 'number', value: 92 })
		expect(s.cells.get(3, 1)?.value).toEqual({ kind: 'number', value: 78 })
	})

	test('boolean values', () => {
		const s = sheetAt(0)
		expect(s.cells.get(1, 2)?.value).toEqual({ kind: 'boolean', value: true })
		expect(s.cells.get(2, 2)?.value).toEqual({ kind: 'boolean', value: true })
		expect(s.cells.get(3, 2)?.value).toEqual({ kind: 'boolean', value: false })
	})

	test('error values', () => {
		const s = sheetAt(0)
		expect(s.cells.get(4, 0)?.value).toEqual({ kind: 'error', value: '#N/A' })
	})

	test('date serial numbers', () => {
		const s = sheetAt(0)
		expect(s.cells.get(1, 3)?.value).toEqual({ kind: 'number', value: 44927 })
		expect(s.cells.get(2, 3)?.value).toEqual({ kind: 'number', value: 44958 })
	})

	test('formulas: SUM', () => {
		const s = sheetAt(0)
		expect(s.cells.get(1, 4)?.formula).toBe('SUM(B2)')
	})

	test('formulas: IF', () => {
		const s = sheetAt(0)
		expect(s.cells.get(5, 0)?.formula).toBe('IF(B2>80,"Pass","Fail")')
	})

	test('formulas: cross-sheet reference', () => {
		const s = sheetAt(1)
		expect(s.cells.get(0, 0)?.formula).toBe('SUM(Data!B2:B4)')
	})

	test('formulas: VLOOKUP', () => {
		const s = sheetAt(1)
		expect(s.cells.get(1, 0)?.formula).toBe('VLOOKUP("Bob",Data!A2:B4,2,FALSE)')
	})

	test('formulas: simple math', () => {
		const s = sheetAt(1)
		expect(s.cells.get(2, 0)?.formula).toBe('2+3')
	})

	test('shared formulas', () => {
		const s = sheetAt(0)
		const master = s.cells.get(1, 4)
		expect(master?.formulaInfo?.kind).toBe('shared')
		expect(master?.formulaInfo?.isMaster).toBe(true)
		expect(master?.formula).toBe('SUM(B2)')

		const member = s.cells.get(2, 4)
		expect(member?.formulaInfo?.kind).toBe('shared')
		expect(member?.formulaInfo?.isMaster).toBe(false)
	})

	test('styles: bold', () => {
		const s = sheetAt(2)
		const cell = must(s.cells.get(0, 0), 'Expected styles!A1')
		const style = result.styles.get(cell.styleId)
		expect(style?.font?.bold).toBe(true)
	})

	test('styles: italic', () => {
		const s = sheetAt(2)
		const cell = must(s.cells.get(0, 1), 'Expected styles!B1')
		const style = result.styles.get(cell.styleId)
		expect(style?.font?.italic).toBe(true)
	})

	test('styles: font color', () => {
		const s = sheetAt(2)
		const cell = must(s.cells.get(0, 2), 'Expected styles!C1')
		const style = result.styles.get(cell.styleId)
		expect(style?.font?.color).toEqual({ kind: 'rgb', rgb: 'FFFF0000' })
	})

	test('styles: number format', () => {
		const s = sheetAt(2)
		const cell = must(s.cells.get(0, 3), 'Expected styles!D1')
		const style = result.styles.get(cell.styleId)
		expect(style?.numberFormat).toBe('0.00%')
	})

	test('styles: background fill', () => {
		const s = sheetAt(2)
		const cell = must(s.cells.get(1, 0), 'Expected styles!A2')
		const style = result.styles.get(cell.styleId)
		expect(style?.fill?.pattern).toBe('solid')
		expect(style?.fill?.fgColor).toEqual({ kind: 'rgb', rgb: 'FF00FF00' })
	})

	test('styles: borders', () => {
		const s = sheetAt(2)
		const cell = must(s.cells.get(1, 1), 'Expected styles!B2')
		const style = result.styles.get(cell.styleId)
		expect(style?.border?.top?.style).toBe('thin')
		expect(style?.border?.bottom?.style).toBe('thin')
	})

	test('styles: alignment', () => {
		const s = sheetAt(2)
		const cell = must(s.cells.get(1, 2), 'Expected styles!C2')
		const style = result.styles.get(cell.styleId)
		expect(style?.alignment?.horizontal).toBe('center')
		expect(style?.alignment?.vertical).toBe('top')
		expect(style?.alignment?.wrapText).toBe(true)
	})

	test('merged cells', () => {
		const s = sheetAt(0)
		expect(s.merges).toHaveLength(2)
		expect(s.merges[0]).toEqual({ start: { row: 6, col: 0 }, end: { row: 6, col: 2 } })
		expect(s.merges[1]).toEqual({ start: { row: 7, col: 0 }, end: { row: 8, col: 0 } })
	})

	test('frozen panes', () => {
		const s = sheetAt(0)
		expect(s.frozenRows).toBe(1)
		expect(s.frozenCols).toBe(1)
	})

	test('column widths', () => {
		const s = sheetAt(0)
		expect(s.colWidths.get(0)).toBe(20)
		expect(s.colWidths.get(1)).toBe(15)
	})

	test('row heights', () => {
		const s = sheetAt(0)
		expect(s.rowHeights.get(0)).toBe(30)
		expect(s.rowHeights.get(5)).toBe(25)
	})

	test('hidden rows', () => {
		const s = sheetAt(0)
		expect(s.rowDefs.get(4)?.hidden).toBe(true)
	})

	test('hidden columns', () => {
		const s = sheetAt(0)
		const hiddenCol = s.colDefs.find((d) => d.min === 3 && d.max === 3)
		expect(hiddenCol?.hidden).toBe(true)
	})

	test('conditional formatting: cellIs greaterThan', () => {
		const s = sheetAt(0)
		const cf = s.conditionalFormats.find((cf) =>
			cf.rules.some((r) => r.type === 'cellIs' && r.operator === 'greaterThan'),
		)
		const foundCf = must(cf, 'Expected greaterThan conditional format')
		expect(foundCf.sqref).toBe('B2:B4')
		const rule = must(
			foundCf.rules.find((r) => r.operator === 'greaterThan'),
			'Expected greaterThan rule',
		)
		expect(rule.formulas).toContain('80')
		expect(rule.style?.font?.bold).toBe(true)
	})

	test('conditional formatting: cellIs lessThan', () => {
		const s = sheetAt(0)
		const cf = s.conditionalFormats.find((cf) =>
			cf.rules.some((r) => r.type === 'cellIs' && r.operator === 'lessThan'),
		)
		const foundCf = must(cf, 'Expected lessThan conditional format')
		const rule = must(
			foundCf.rules.find((r) => r.operator === 'lessThan'),
			'Expected lessThan rule',
		)
		expect(rule.formulas).toContain('80')
		expect(rule.style?.font?.italic).toBe(true)
	})

	test('data validation: dropdown list', () => {
		const s = sheetAt(0)
		const dv = s.dataValidations.find((d) => d.type === 'list')
		const found = must(dv, 'Expected list validation')
		expect(found.sqref).toBe('C2:C4')
		expect(found.formula1).toBe('"Yes,No,Maybe"')
	})

	test('data validation: number range', () => {
		const s = sheetAt(0)
		const dv = s.dataValidations.find((d) => d.type === 'whole')
		const found = must(dv, 'Expected whole-number validation')
		expect(found.sqref).toBe('B2:B4')
		expect(found.operator).toBe('between')
		expect(found.formula1).toBe('0')
		expect(found.formula2).toBe('100')
	})

	test('comments', () => {
		const s = sheetAt(0)
		expect(s.comments.get('A1')).toMatchObject({
			text: 'Student name',
			author: 'Teacher',
			legacyDrawing: {
				row: 0,
				column: 0,
				visible: false,
			},
		})
		expect(s.comments.get('B1')?.text).toBe('Exam score out of 100')
	})

	test('hyperlinks', () => {
		const data = sheetAt(0)
		expect(data.hyperlinks.get('A1')?.target).toBe('https://example.com/students')
		expect(data.hyperlinks.get('A1')?.display).toBe('Name')
		expect(data.hyperlinks.get('A1')?.tooltip).toBe('View students')

		const summary = sheetAt(1)
		expect(summary.hyperlinks.get('A1')?.target).toBe('https://example.com/summary')
	})

	test('defined names', () => {
		expect(result.definedNames.get('Scores')).toBe('Data!$B$2:$B$4')
		expect(result.definedNames.get('StudentNames')).toBe('Data!$A$2:$A$4')
	})

	test('auto-filter', () => {
		const s = sheetAt(0)
		const filter = must(s.autoFilter, 'Expected auto filter')
		expect(filter.ref).toBe('A1:E4')
	})

	test('tables', () => {
		const s = sheetAt(3)
		expect(s.tables).toHaveLength(1)
		const table = must(s.tables[0], 'Expected table')
		expect(table.name).toBe('Products')
		expect(table.hasHeaders).toBe(true)
		expect(table.ref).toEqual({ start: { row: 0, col: 0 }, end: { row: 2, col: 1 } })
		expect(table.columns).toHaveLength(2)
		expect(table.columns[0]?.name).toBe('Product')
		expect(table.columns[1]?.name).toBe('Price')
	})

	test('sheet protection', () => {
		const s = sheetAt(1)
		const protection = must(s.protection, 'Expected sheet protection')
		expect(protection.sheet).toBe(true)
		expect(protection.objects).toBe(true)
		expect(protection.scenarios).toBe(true)
		expect(protection.sort).toBe(false)
		expect(protection.autoFilter).toBe(false)
	})
})
