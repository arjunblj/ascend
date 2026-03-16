import { describe, expect, test } from 'bun:test'
import { createWorkbook } from '@ascend/core'
import { EMPTY, numberValue, stringValue } from '@ascend/schema'
import { validateCellValue } from './data-validation.ts'

describe('validateCellValue', () => {
	test('cells without validation pass', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		expect(validateCellValue(sheet, 0, 0, numberValue(42), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, stringValue('hello'), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, EMPTY, wb)).toEqual({ valid: true })
	})

	test('whole number validation - value between 1 and 10', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'A1:A5',
			type: 'whole',
			formula1: '1',
			formula2: '10',
			operator: 'between',
		})

		expect(validateCellValue(sheet, 0, 0, numberValue(5), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, numberValue(1), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, numberValue(10), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, numberValue(0), wb).valid).toBe(false)
		expect(validateCellValue(sheet, 0, 0, numberValue(11), wb).valid).toBe(false)
		expect(validateCellValue(sheet, 0, 0, numberValue(3.5), wb).valid).toBe(false)
	})

	test('whole number validation - cell outside range has no rule', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'A1:A5',
			type: 'whole',
			formula1: '1',
			formula2: '10',
		})
		expect(validateCellValue(sheet, 0, 1, numberValue(999), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 5, 0, numberValue(999), wb)).toEqual({ valid: true })
	})

	test('list validation - value must be one of allowed values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'B1:B3',
			type: 'list',
			formula1: '"Red,Green,Blue"',
		})

		expect(validateCellValue(sheet, 0, 1, stringValue('Red'), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 1, stringValue('Green'), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 1, stringValue('Blue'), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 1, stringValue('Yellow'), wb).valid).toBe(false)
		expect(validateCellValue(sheet, 0, 1, stringValue('red'), wb).valid).toBe(false)
	})

	test('list validation - unquoted comma-separated values', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'A1',
			type: 'list',
			formula1: 'A,B,C',
		})
		expect(validateCellValue(sheet, 0, 0, stringValue('A'), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, stringValue('B'), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, stringValue('D'), wb).valid).toBe(false)
	})

	test('text length validation - max 5 characters', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'C1:C10',
			type: 'textLength',
			formula1: '0',
			formula2: '5',
			operator: 'between',
		})

		expect(validateCellValue(sheet, 0, 2, stringValue('hi'), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 2, stringValue('hello'), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 2, stringValue(''), wb).valid).toBe(true)
		expect(validateCellValue(sheet, 0, 2, stringValue('hello!'), wb).valid).toBe(false)
	})

	test('text length validation - returns custom error message when invalid', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'A1',
			type: 'textLength',
			formula1: '0',
			formula2: '3',
			error: 'Max 3 chars allowed',
		})
		const result = validateCellValue(sheet, 0, 0, stringValue('abcd'), wb)
		expect(result.valid).toBe(false)
		expect(result.message).toBe('Max 3 chars allowed')
		expect(result.rule).toBeDefined()
	})

	test('allowBlank - empty passes when allowBlank is true', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'A1',
			type: 'whole',
			formula1: '1',
			formula2: '10',
			allowBlank: true,
		})
		expect(validateCellValue(sheet, 0, 0, EMPTY, wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, stringValue(''), wb)).toEqual({ valid: true })
	})

	test('allowBlank - empty fails when allowBlank is false', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'A1',
			type: 'list',
			formula1: '"A,B"',
			allowBlank: false,
		})
		const result = validateCellValue(sheet, 0, 0, EMPTY, wb)
		expect(result.valid).toBe(false)
		expect(result.message).toBeDefined()
	})

	test('decimal validation - accepts numbers within range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'A1',
			type: 'decimal',
			formula1: '0',
			formula2: '100',
			operator: 'between',
		})
		expect(validateCellValue(sheet, 0, 0, numberValue(50), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, numberValue(0.5), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, numberValue(101), wb).valid).toBe(false)
	})

	test('date validation - accepts date serial within range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'A1',
			type: 'date',
			formula1: '44927',
			formula2: '44957',
			operator: 'between',
		})
		expect(validateCellValue(sheet, 0, 0, { kind: 'date', serial: 44940 }, wb)).toEqual({
			valid: true,
		})
		expect(validateCellValue(sheet, 0, 0, { kind: 'date', serial: 44900 }, wb).valid).toBe(false)
	})

	test('custom validation evaluates formula result', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.dataValidations.push({
			sqref: 'A1',
			type: 'custom',
			formula1: '=A1>0',
		})
		expect(validateCellValue(sheet, 0, 0, numberValue(1), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, numberValue(0), wb).valid).toBe(false)
	})

	test('list validation can resolve a referenced range', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.cells.set(0, 1, { value: stringValue('Red'), formula: null, styleId: 0 as never })
		sheet.cells.set(1, 1, { value: stringValue('Blue'), formula: null, styleId: 0 as never })
		sheet.dataValidations.push({
			sqref: 'A1',
			type: 'list',
			formula1: 'B1:B2',
		})
		expect(validateCellValue(sheet, 0, 0, stringValue('Red'), wb)).toEqual({ valid: true })
		expect(validateCellValue(sheet, 0, 0, stringValue('Green'), wb).valid).toBe(false)
	})
})
