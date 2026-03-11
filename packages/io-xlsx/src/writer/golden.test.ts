import { describe, expect, it } from 'bun:test'
import type { StyleId } from '@ascend/core'
import { createTableId, Workbook } from '@ascend/core'
import { numberValue, stringValue } from '@ascend/schema'
import { fingerprintXlsx } from '../../test/fidelity-harness.ts'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

const S0 = 0 as StyleId

function roundTrip(wb: Workbook) {
	const written = writeXlsx(wb)
	if (!written.ok) throw new Error(`write failed: ${written.error.message}`)
	const read = readXlsx(written.value)
	if (!read.ok) throw new Error(`read failed: ${read.error.message}`)
	return { bytes: written.value, result: read.value }
}

function cellValueEqual(
	a: { kind: string; value?: unknown } | undefined,
	b: { kind: string; value?: unknown } | undefined,
): boolean {
	if (!a || !b) return a === b
	if (a.kind !== b.kind) return false
	if (a.kind === 'number' && b.kind === 'number') {
		const an = a.value as number
		const bn = b.value as number
		if (Number.isNaN(an) && Number.isNaN(bn)) return true
		return an === bn || Math.abs(an - bn) < 1e-10
	}
	return JSON.stringify(a) === JSON.stringify(b)
}

describe('golden round-trip fidelity', () => {
	it('numeric workbook fidelity', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Numeric')
		sheet.cells.set(0, 0, { value: numberValue(0), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(42), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: numberValue(-100), formula: null, styleId: S0 })
		sheet.cells.set(0, 3, { value: numberValue(Math.E), formula: null, styleId: S0 })
		sheet.cells.set(0, 4, { value: numberValue(1e15), formula: null, styleId: S0 })
		sheet.cells.set(0, 5, { value: numberValue(-0.0001), formula: null, styleId: S0 })

		const { bytes: bytes1, result: result1 } = roundTrip(wb)
		const { bytes: bytes2, result: result2 } = roundTrip(result1.workbook)

		const fp1 = fingerprintXlsx(bytes1)
		const fp2 = fingerprintXlsx(bytes2)
		expect(fp1.partPaths).toEqual(fp2.partPaths)
		expect(fp1.workbook?.tagCounts).toMatchObject(fp2.workbook?.tagCounts ?? {})
		expect(fp1.sheets.length).toBe(fp2.sheets.length)

		const s1 = result1.workbook.sheets[0]
		const s2 = result2.workbook.sheets[0]
		expect(s1).toBeDefined()
		expect(s2).toBeDefined()
		for (let col = 0; col <= 5; col++) {
			expect(cellValueEqual(s1?.cells.get(0, col)?.value, s2?.cells.get(0, col)?.value)).toBe(true)
		}
		expect(s2?.cells.get(0, 0)?.value).toEqual({ kind: 'number', value: 0 })
		expect(s2?.cells.get(0, 1)?.value).toEqual({ kind: 'number', value: 42 })
		expect(s2?.cells.get(0, 2)?.value).toEqual({ kind: 'number', value: -100 })
		expect(s2?.cells.get(0, 4)?.value).toEqual({ kind: 'number', value: 1e15 })
	})

	it('formula workbook fidelity', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Formulas')
		sheet.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: S0 })
		sheet.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: S0 })
		sheet.cells.set(0, 2, { value: numberValue(30), formula: null, styleId: S0 })
		sheet.cells.set(1, 0, { value: numberValue(60), formula: 'SUM(A1:C1)', styleId: S0 })
		sheet.cells.set(1, 1, { value: numberValue(1), formula: 'IF(A1>5,1,0)', styleId: S0 })
		sheet.cells.set(0, 3, { value: numberValue(100), formula: null, styleId: S0 })
		sheet.cells.set(1, 3, { value: numberValue(200), formula: null, styleId: S0 })
		sheet.cells.set(2, 0, {
			value: numberValue(100),
			formula: 'VLOOKUP(A1,A1:D2,4,FALSE)',
			styleId: S0,
		})

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		expect(s?.cells.get(1, 0)?.formula).toBe('SUM(A1:C1)')
		expect(s?.cells.get(1, 1)?.formula).toBe('IF(A1>5,1,0)')
		expect(s?.cells.get(2, 0)?.formula).toBe('VLOOKUP(A1,A1:D2,4,FALSE)')
	})

	it('styled workbook fidelity', () => {
		const wb = new Workbook()
		const boldId = wb.styles.register({ font: { bold: true } })
		const colorId = wb.styles.register({
			font: { color: { kind: 'rgb', rgb: 'FFFF0000' } },
		})
		const pctId = wb.styles.register({ numberFormat: '0.0%' })
		const sheet = wb.addSheet('Styled')
		sheet.cells.set(0, 0, { value: stringValue('Bold'), formula: null, styleId: boldId })
		sheet.cells.set(0, 1, { value: stringValue('Red'), formula: null, styleId: colorId })
		sheet.cells.set(0, 2, { value: numberValue(0.75), formula: null, styleId: pctId })

		const { result } = roundTrip(wb)
		const s = result.workbook.sheets[0]
		const cell0 = s?.cells.get(0, 0)
		const cell1 = s?.cells.get(0, 1)
		const cell2 = s?.cells.get(0, 2)
		expect(result.workbook.styles.get(cell0?.styleId ?? (0 as StyleId))?.font?.bold).toBe(true)
		expect(result.workbook.styles.get(cell1?.styleId ?? (0 as StyleId))?.font?.color).toEqual({
			kind: 'rgb',
			rgb: 'FFFF0000',
		})
		expect(result.workbook.styles.get(cell2?.styleId ?? (0 as StyleId))?.numberFormat).toBe('0.0%')
	})

	it('multi-sheet fidelity', () => {
		const wb = new Workbook()
		const s1 = wb.addSheet('Input')
		s1.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: S0 })
		s1.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: S0 })
		const s2 = wb.addSheet('Calc')
		s2.cells.set(0, 0, { value: numberValue(3), formula: 'Input!A1+Input!B1', styleId: S0 })
		wb.definedNames.set('Total', 'Calc!$A$1')
		const s3 = wb.addSheet('TableSheet')
		s3.cells.set(0, 0, { value: stringValue('Name'), formula: null, styleId: S0 })
		s3.cells.set(0, 1, { value: stringValue('Qty'), formula: null, styleId: S0 })
		s3.cells.set(1, 0, { value: stringValue('Item1'), formula: null, styleId: S0 })
		s3.cells.set(1, 1, { value: numberValue(5), formula: null, styleId: S0 })
		s3.tables.push({
			id: createTableId(),
			name: 'DataTable',
			sheetId: s3.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Qty' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const { result, bytes } = roundTrip(wb)
		expect(result.workbook.sheets).toHaveLength(3)
		expect(result.workbook.sheets[0]?.name).toBe('Input')
		expect(result.workbook.sheets[1]?.name).toBe('Calc')
		expect(result.workbook.sheets[2]?.name).toBe('TableSheet')
		expect(result.workbook.sheets[1]?.cells.get(0, 0)?.formula).toBe('Input!A1+Input!B1')
		expect(result.workbook.definedNames.get('Total')).toBe('Calc!$A$1')
		expect(result.workbook.sheets[2]?.tables).toHaveLength(1)
		expect(result.workbook.sheets[2]?.tables[0]?.name).toBe('DataTable')

		const fp = fingerprintXlsx(bytes)
		expect(fp.sheets).toHaveLength(3)
		expect(fp.partPaths).toContain('xl/workbook.xml')
		expect(fp.workbook?.tagCounts?.sheet).toBe(3)
		expect(fp.workbook?.tagCounts?.definedName).toBeGreaterThanOrEqual(1)
	})
})
