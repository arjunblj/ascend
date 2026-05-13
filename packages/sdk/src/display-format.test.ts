import { describe, expect, test } from 'bun:test'
import { AscendWorkbook, formatStyledDisplayCellValue } from './index.ts'

describe('style-aware display formatting', () => {
	test('formats scalar values with common Excel number formats', () => {
		expect(formatStyledDisplayCellValue({ kind: 'number', value: 1234.5 })).toBe('1234.5')
		expect(
			formatStyledDisplayCellValue({ kind: 'number', value: 1234.5 }, { numberFormat: '#,##0.00' }),
		).toBe('1,234.50')
		expect(
			formatStyledDisplayCellValue({ kind: 'number', value: 0.125 }, { numberFormat: '0.00%' }),
		).toBe('12.50%')
		expect(
			formatStyledDisplayCellValue({ kind: 'number', value: 999 }, { numberFormat: '$#,##0' }),
		).toBe('$999')
		expect(
			formatStyledDisplayCellValue(
				{ kind: 'number', value: -100 },
				{ numberFormat: '#,##0.00;[Red]-#,##0.00' },
			),
		).toBe('-100.00')
	})

	test('formats date serials with simple date formats and workbook date systems', () => {
		expect(
			formatStyledDisplayCellValue({ kind: 'date', serial: 45292 }, { numberFormat: 'yyyy-mm-dd' }),
		).toBe('2024-01-01')
		expect(
			formatStyledDisplayCellValue({ kind: 'date', serial: 44927 }, { numberFormat: 'dd/mm/yyyy' }),
		).toBe('01/01/2023')
		expect(
			formatStyledDisplayCellValue(
				{ kind: 'date', serial: 0 },
				{ numberFormat: 'm/d/yy' },
				{ dateSystem: '1904' },
			),
		).toBe('1/1/04')
	})

	test('leaves non-numeric values semantic while honoring text and error display', () => {
		expect(
			formatStyledDisplayCellValue({ kind: 'string', value: 'ID-001' }, { numberFormat: '@' }),
		).toBe('ID-001')
		expect(
			formatStyledDisplayCellValue({ kind: 'boolean', value: true }, { numberFormat: '0%' }),
		).toBe('TRUE')
		expect(
			formatStyledDisplayCellValue({ kind: 'error', value: '#DIV/0!' }, { numberFormat: '0%' }),
		).toBe('#DIV/0!')
	})

	test('workbook read views format stored and calculated cell values without changing raw reads', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 0.25 }] },
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'A1:A1', format: '0.0%' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'B1:B1', format: '0.00%' },
		])
		wb.recalc()

		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 0.25 })
		expect(wb.formatCellForDisplay('Sheet1!A1')).toBe('25.0%')
		expect(wb.formatCellForDisplay('Sheet1!B1')).toBe('50.00%')
		expect(wb.sheet('Sheet1')?.formatCellForDisplay('B1')).toBe('50.00%')
	})
})
