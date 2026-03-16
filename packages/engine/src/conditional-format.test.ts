import { describe, expect, test } from 'bun:test'
import type { CellStyle } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import { EMPTY, numberValue, stringValue } from '@ascend/schema'
import { evaluateConditionalFormats } from './conditional-format.ts'

const sid = 0
const greenFill: CellStyle = {
	fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
}
const boldStyle: CellStyle = { font: { bold: true } }

function setCell(
	sheet: ReturnType<ReturnType<typeof createWorkbook>['addSheet']>,
	row: number,
	col: number,
	value: ReturnType<typeof numberValue> | ReturnType<typeof stringValue> | typeof EMPTY,
) {
	sheet.cells.set(row, col, { value, formula: null, styleId: sid })
}

describe('evaluateConditionalFormats', () => {
	test('cellIs greaterThan rule matches correct cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, numberValue(1))
		setCell(sheet, 1, 0, numberValue(5))
		setCell(sheet, 2, 0, numberValue(10))
		setCell(sheet, 3, 0, numberValue(4))

		sheet.conditionalFormats.push({
			sqref: 'A1:A5',
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['3'],
					priority: 1,
					style: greenFill,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)

		expect(result.get('A1')).toBeUndefined()
		expect(result.get('A2')).toBeDefined()
		expect(result.get('A2')?.[0]).toMatchObject({
			ruleIndex: 0,
			priority: 1,
			type: 'cellIs',
			format: greenFill,
		})
		expect(result.get('A3')).toBeDefined()
		expect(result.get('A4')).toBeDefined()
		expect(result.get('A5')).toBeUndefined()
	})

	test('containsText rule matches correct cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, stringValue('hello world'))
		setCell(sheet, 1, 0, stringValue('foo bar'))
		setCell(sheet, 2, 0, stringValue('hello'))
		setCell(sheet, 3, 0, stringValue('goodbye'))

		sheet.conditionalFormats.push({
			sqref: 'A1:A4',
			rules: [
				{
					type: 'containsText',
					formulas: ['"hello"'],
					priority: 1,
					style: boldStyle,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)

		expect(result.get('A1')).toBeDefined()
		expect(result.get('A2')).toBeUndefined()
		expect(result.get('A3')).toBeDefined()
		expect(result.get('A4')).toBeUndefined()
	})

	test('containsBlanks rule matches empty cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, numberValue(1))
		setCell(sheet, 1, 0, EMPTY)
		setCell(sheet, 2, 0, stringValue(''))
		setCell(sheet, 3, 0, stringValue('x'))

		sheet.conditionalFormats.push({
			sqref: 'A1:A4',
			rules: [
				{
					type: 'containsBlanks',
					formulas: [],
					priority: 1,
					style: greenFill,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)

		expect(result.get('A1')).toBeUndefined()
		expect(result.get('A2')).toBeDefined()
		expect(result.get('A2')?.[0]).toMatchObject({ type: 'containsBlanks' })
		expect(result.get('A3')).toBeDefined()
		expect(result.get('A4')).toBeUndefined()
	})

	test('notContainsBlanks rule matches non-empty cells', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, EMPTY)
		setCell(sheet, 1, 0, numberValue(42))
		setCell(sheet, 2, 0, stringValue('hello'))

		sheet.conditionalFormats.push({
			sqref: 'A1:A3',
			rules: [
				{
					type: 'notContainsBlanks',
					formulas: [],
					priority: 1,
					style: boldStyle,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)

		expect(result.get('A1')).toBeUndefined()
		expect(result.get('A2')).toBeDefined()
		expect(result.get('A3')).toBeDefined()
	})

	test('beginsWith and endsWith rules', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, stringValue('hello'))
		setCell(sheet, 1, 0, stringValue('world'))
		setCell(sheet, 2, 0, stringValue('hello world'))

		sheet.conditionalFormats.push({
			sqref: 'A1:A3',
			rules: [
				{
					type: 'beginsWith',
					formulas: ['"hel"'],
					priority: 1,
					style: boldStyle,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)

		expect(result.get('A1')).toBeDefined()
		expect(result.get('A2')).toBeUndefined()
		expect(result.get('A3')).toBeDefined()
	})

	test('cellIs lessThan and between', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, numberValue(1))
		setCell(sheet, 1, 0, numberValue(5))
		setCell(sheet, 2, 0, numberValue(10))

		sheet.conditionalFormats.push({
			sqref: 'A1:A3',
			rules: [
				{
					type: 'cellIs',
					operator: 'between',
					formulas: ['3', '8'],
					priority: 1,
					style: greenFill,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)

		expect(result.get('A1')).toBeUndefined()
		expect(result.get('A2')).toBeDefined()
		expect(result.get('A3')).toBeUndefined()
	})

	test('cells outside sqref are not evaluated', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, numberValue(100))
		setCell(sheet, 0, 1, numberValue(100))

		sheet.conditionalFormats.push({
			sqref: 'A1',
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['50'],
					priority: 1,
					style: greenFill,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)

		expect(result.get('A1')).toBeDefined()
		expect(result.get('B1')).toBeUndefined()
	})

	test('formula rules evaluate relative to the top-left cell of sqref', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, numberValue(1))
		setCell(sheet, 1, 0, numberValue(5))
		setCell(sheet, 2, 0, numberValue(10))

		sheet.conditionalFormats.push({
			sqref: 'A1:A3',
			rules: [
				{
					type: 'expression',
					formulas: ['A1>3'],
					priority: 1,
					style: greenFill,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)
		expect(result.get('A1')).toBeUndefined()
		expect(result.get('A2')).toBeDefined()
		expect(result.get('A3')).toBeDefined()
	})
})
