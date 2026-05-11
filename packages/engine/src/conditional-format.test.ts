import { describe, expect, test } from 'bun:test'
import type { CellStyle, SheetConditionalFormat } from '@ascend/core'
import { createWorkbook } from '@ascend/core'
import type { CellValue } from '@ascend/schema'
import {
	booleanValue,
	dateValue,
	EMPTY,
	errorValue,
	numberValue,
	stringValue,
} from '@ascend/schema'
import { evaluateConditionalFormats } from './conditional-format.ts'

const sid = 0
const greenFill: CellStyle = {
	fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFC6EFCE' } },
}
const redFill: CellStyle = {
	fill: { pattern: 'solid', fgColor: { kind: 'rgb', rgb: 'FFFFC7CE' } },
}
const boldStyle: CellStyle = { font: { bold: true } }

function setCell(
	sheet: ReturnType<ReturnType<typeof createWorkbook>['addSheet']>,
	row: number,
	col: number,
	value: CellValue,
) {
	sheet.cells.set(row, col, { value, formula: null, styleId: sid })
}

function makeRule(
	type: string,
	opts: {
		operator?: string
		formulas?: string[]
		priority?: number
		style?: CellStyle
		stopIfTrue?: boolean
		rank?: number
		percent?: boolean
		bottom?: boolean
		aboveAverage?: boolean
		equalAverage?: boolean
		stdDev?: number
		text?: string
		timePeriod?: string
	} = {},
) {
	return {
		type,
		operator: opts.operator,
		formulas: opts.formulas ?? [],
		priority: opts.priority ?? 1,
		style: opts.style ?? greenFill,
		stopIfTrue: opts.stopIfTrue,
		...(opts.rank !== undefined ? { rank: opts.rank } : {}),
		...(opts.percent !== undefined ? { percent: opts.percent } : {}),
		...(opts.bottom !== undefined ? { bottom: opts.bottom } : {}),
		...(opts.aboveAverage !== undefined ? { aboveAverage: opts.aboveAverage } : {}),
		...(opts.equalAverage !== undefined ? { equalAverage: opts.equalAverage } : {}),
		...(opts.stdDev !== undefined ? { stdDev: opts.stdDev } : {}),
		...(opts.text !== undefined ? { text: opts.text } : {}),
		...(opts.timePeriod !== undefined ? { timePeriod: opts.timePeriod } : {}),
	}
}

function makeCf(sqref: string, ...rules: ReturnType<typeof makeRule>[]): SheetConditionalFormat {
	return { sqref, rules }
}

describe('evaluateConditionalFormats', () => {
	// ── cellIs operator tests ─────────────────────────────────────────

	describe('cellIs', () => {
		test('equal', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(5))
			setCell(sheet, 1, 0, numberValue(3))
			setCell(sheet, 2, 0, numberValue(5))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('cellIs', { operator: 'equal', formulas: ['5'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
			expect(result.get('A3')).toBeDefined()
		})

		test('notEqual', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(5))
			setCell(sheet, 1, 0, numberValue(3))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('cellIs', { operator: 'notEqual', formulas: ['5'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
		})

		test('text equality matches string cells', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('Closed'))
			setCell(sheet, 1, 0, stringValue('closed'))
			setCell(sheet, 2, 0, stringValue('Open'))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('cellIs', { operator: 'equal', formulas: ['"Closed"'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('text comparison can use a relative reference criterion', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('North'))
			setCell(sheet, 1, 0, stringValue('South'))
			setCell(sheet, 0, 1, stringValue('north'))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('cellIs', { operator: 'equal', formulas: ['$B$1'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})

		test('text notEqual matches cells with different text', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('Closed'))
			setCell(sheet, 1, 0, stringValue('Open'))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('cellIs', { operator: 'notEqual', formulas: ['"Closed"'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
		})

		test('greaterThan', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(1))
			setCell(sheet, 1, 0, numberValue(5))
			setCell(sheet, 2, 0, numberValue(10))
			setCell(sheet, 3, 0, numberValue(4))

			sheet.conditionalFormats.push(
				makeCf('A1:A5', makeRule('cellIs', { operator: 'greaterThan', formulas: ['3'] })),
			)

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

		test('lessThan', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(1))
			setCell(sheet, 1, 0, numberValue(5))
			setCell(sheet, 2, 0, numberValue(10))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('cellIs', { operator: 'lessThan', formulas: ['5'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('greaterThanOrEqual', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(4))
			setCell(sheet, 1, 0, numberValue(5))
			setCell(sheet, 2, 0, numberValue(6))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('cellIs', { operator: 'greaterThanOrEqual', formulas: ['5'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeDefined()
		})

		test('lessThanOrEqual', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(4))
			setCell(sheet, 1, 0, numberValue(5))
			setCell(sheet, 2, 0, numberValue(6))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('cellIs', { operator: 'lessThanOrEqual', formulas: ['5'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('between (inclusive)', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(1))
			setCell(sheet, 1, 0, numberValue(3))
			setCell(sheet, 2, 0, numberValue(5))
			setCell(sheet, 3, 0, numberValue(8))
			setCell(sheet, 4, 0, numberValue(10))

			sheet.conditionalFormats.push(
				makeCf('A1:A5', makeRule('cellIs', { operator: 'between', formulas: ['3', '8'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined() // 3 is inclusive lower bound
			expect(result.get('A3')).toBeDefined() // 5 is in range
			expect(result.get('A4')).toBeDefined() // 8 is inclusive upper bound
			expect(result.get('A5')).toBeUndefined()
		})

		test('notBetween', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(1))
			setCell(sheet, 1, 0, numberValue(5))
			setCell(sheet, 2, 0, numberValue(10))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('cellIs', { operator: 'notBetween', formulas: ['3', '8'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
			expect(result.get('A3')).toBeDefined()
		})

		test('empty cell never matches cellIs', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, EMPTY)

			sheet.conditionalFormats.push(
				makeCf('A1', makeRule('cellIs', { operator: 'equal', formulas: ['0'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
		})

		test('cellIs with formula reference', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 0, 1, numberValue(10)) // B1 = 10 used as threshold

			sheet.conditionalFormats.push(
				makeCf('A1', makeRule('cellIs', { operator: 'greaterThanOrEqual', formulas: ['B1'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
		})
	})

	// ── expression (custom formula) ───────────────────────────────────

	describe('expression', () => {
		test('simple boolean formula', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(1))
			setCell(sheet, 1, 0, numberValue(5))
			setCell(sheet, 2, 0, numberValue(10))

			sheet.conditionalFormats.push(makeCf('A1:A3', makeRule('expression', { formulas: ['A1>3'] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeDefined()
		})

		test('formula referencing another column', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 0, 1, numberValue(5))
			setCell(sheet, 1, 0, numberValue(3))
			setCell(sheet, 1, 1, numberValue(7))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('expression', { formulas: ['A1>B1'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined() // 10 > 5
			expect(result.get('A2')).toBeUndefined() // 3 < 7
		})

		test('formula with no formula returns false', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(1))

			sheet.conditionalFormats.push(makeCf('A1', makeRule('expression', { formulas: [] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
		})

		test('formula returning numeric truthy/falsy', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(0))
			setCell(sheet, 1, 0, numberValue(42))

			sheet.conditionalFormats.push(makeCf('A1:A2', makeRule('expression', { formulas: ['A1'] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined() // 0 is falsy
			expect(result.get('A2')).toBeDefined() // 42 is truthy
		})
	})

	// ── colorScale (visual rule — structure round-trip) ───────────────

	describe('colorScale', () => {
		test('2-color scale rule structure is stored and retrievable', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 1, 0, numberValue(50))
			setCell(sheet, 2, 0, numberValue(90))

			const rule = {
				type: 'colorScale' as const,
				formulas: [] as string[],
				priority: 1,
			}
			sheet.conditionalFormats.push({ sqref: 'A1:A3', rules: [rule] })

			expect(sheet.conditionalFormats).toHaveLength(1)
			expect(sheet.conditionalFormats[0].rules[0].type).toBe('colorScale')
			expect(sheet.conditionalFormats[0].sqref).toBe('A1:A3')
		})

		test('3-color scale rule round-trips through clone', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')

			const rule = {
				type: 'colorScale' as const,
				formulas: [] as string[],
				priority: 1,
			}
			sheet.conditionalFormats.push({ sqref: 'A1:A10', rules: [rule] })

			const cloned = sheet.clone()
			expect(cloned.conditionalFormats).toHaveLength(1)
			expect(cloned.conditionalFormats[0].rules[0].type).toBe('colorScale')
			expect(cloned.conditionalFormats[0].sqref).toBe('A1:A10')
		})
	})

	// ── dataBar (visual rule — structure round-trip) ──────────────────

	describe('dataBar', () => {
		test('dataBar rule structure is stored', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 1, 0, numberValue(50))
			setCell(sheet, 2, 0, numberValue(90))

			const rule = {
				type: 'dataBar' as const,
				formulas: [] as string[],
				priority: 1,
			}
			sheet.conditionalFormats.push({ sqref: 'A1:A3', rules: [rule] })

			expect(sheet.conditionalFormats[0].rules[0].type).toBe('dataBar')
		})

		test('dataBar with negative values is stored correctly', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(-20))
			setCell(sheet, 1, 0, numberValue(0))
			setCell(sheet, 2, 0, numberValue(50))

			const rule = {
				type: 'dataBar' as const,
				formulas: [] as string[],
				priority: 1,
			}
			sheet.conditionalFormats.push({ sqref: 'A1:A3', rules: [rule] })

			const cloned = sheet.clone()
			expect(cloned.conditionalFormats[0].rules[0].type).toBe('dataBar')
		})
	})

	// ── iconSet (visual rule — structure round-trip) ──────────────────

	describe('iconSet', () => {
		test('3-icon set rule structure round-trips', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 1, 0, numberValue(50))
			setCell(sheet, 2, 0, numberValue(90))

			const rule = {
				type: 'iconSet' as const,
				formulas: [] as string[],
				priority: 1,
			}
			sheet.conditionalFormats.push({ sqref: 'A1:A3', rules: [rule] })

			const cloned = sheet.clone()
			expect(cloned.conditionalFormats[0].rules[0].type).toBe('iconSet')
		})

		test('5-icon set rule structure is stored', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			for (let i = 0; i < 5; i++) {
				setCell(sheet, i, 0, numberValue(i * 25))
			}

			const rule = {
				type: 'iconSet' as const,
				formulas: [] as string[],
				priority: 1,
			}
			sheet.conditionalFormats.push({ sqref: 'A1:A5', rules: [rule] })

			expect(sheet.conditionalFormats[0].rules[0].type).toBe('iconSet')
			expect(sheet.conditionalFormats[0].sqref).toBe('A1:A5')
		})
	})

	// ── top10 ────────────────────────────────────────────────────────

	describe('top10', () => {
		test('top 3 values match', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(100))
			setCell(sheet, 1, 0, numberValue(80))
			setCell(sheet, 2, 0, numberValue(60))
			setCell(sheet, 3, 0, numberValue(40))
			setCell(sheet, 4, 0, numberValue(20))

			sheet.conditionalFormats.push(
				makeCf('A1:A5', makeRule('top10', { rank: 3, style: greenFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(3)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeDefined()
			expect(result.get('A4')).toBeUndefined()
			expect(result.get('A5')).toBeUndefined()
		})

		test('bottom 3 values match', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			for (let i = 0; i < 10; i++) setCell(sheet, i, 0, numberValue((i + 1) * 10))

			sheet.conditionalFormats.push(
				makeCf('A1:A10', makeRule('top10', { rank: 3, bottom: true, style: redFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(3)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeDefined()
			expect(result.get('A4')).toBeUndefined()
		})

		test('top 10 percent mode matches top 2 of 20', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			for (let i = 0; i < 20; i++) setCell(sheet, i, 0, numberValue(i + 1))

			sheet.conditionalFormats.push(
				makeCf('A1:A20', makeRule('top10', { rank: 10, percent: true, style: greenFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
			expect(result.get('A19')).toBeDefined()
			expect(result.get('A20')).toBeDefined()
			expect(result.get('A18')).toBeUndefined()
		})

		test('non-numeric cells are skipped', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 1, 0, stringValue('hello'))
			setCell(sheet, 2, 0, numberValue(20))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('top10', { rank: 1, style: greenFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A3')).toBeDefined()
		})

		test('numeric-looking text and booleans are not ranked as numbers', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(100))
			setCell(sheet, 1, 0, stringValue('200'))
			setCell(sheet, 2, 0, booleanValue(true))
			setCell(sheet, 3, 0, numberValue(50))

			sheet.conditionalFormats.push(
				makeCf('A1:A4', makeRule('top10', { rank: 1, style: greenFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('date cells rank by serial value', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, dateValue(45000))
			setCell(sheet, 1, 0, numberValue(100))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('top10', { rank: 1, style: greenFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})
	})

	// ── aboveAverage ─────────────────────────────────────────────────

	describe('aboveAverage', () => {
		test('above average (strictly greater)', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 1, 0, numberValue(20))
			setCell(sheet, 2, 0, numberValue(30))
			setCell(sheet, 3, 0, numberValue(40))
			setCell(sheet, 4, 0, numberValue(50))

			sheet.conditionalFormats.push(makeCf('A1:A5', makeRule('aboveAverage', { style: greenFill })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
			expect(result.get('A4')).toBeDefined()
			expect(result.get('A5')).toBeDefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('below average (strictly less)', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 1, 0, numberValue(20))
			setCell(sheet, 2, 0, numberValue(30))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('aboveAverage', { aboveAverage: false, style: redFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})

		test('equal or above average', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(20))
			setCell(sheet, 1, 0, numberValue(20))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('aboveAverage', { equalAverage: true, style: greenFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
		})

		test('equal or below average', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(20))
			setCell(sheet, 1, 0, numberValue(20))

			sheet.conditionalFormats.push(
				makeCf(
					'A1:A2',
					makeRule('aboveAverage', {
						aboveAverage: false,
						equalAverage: true,
						style: redFill,
					}),
				),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
		})

		test('above one standard deviation', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			for (let row = 0; row < 5; row++) {
				setCell(sheet, row, 0, numberValue((row + 1) * 10))
			}

			sheet.conditionalFormats.push(
				makeCf('A1:A5', makeRule('aboveAverage', { stdDev: 1, style: greenFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A5')).toBeDefined()
			expect(result.get('A4')).toBeUndefined()
		})

		test('below one standard deviation', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			for (let row = 0; row < 5; row++) {
				setCell(sheet, row, 0, numberValue((row + 1) * 10))
			}

			sheet.conditionalFormats.push(
				makeCf(
					'A1:A5',
					makeRule('aboveAverage', { aboveAverage: false, stdDev: 1, style: redFill }),
				),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})

		test('numeric-looking text and booleans do not affect the average', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 1, 0, numberValue(30))
			setCell(sheet, 2, 0, stringValue('1000'))
			setCell(sheet, 3, 0, booleanValue(true))

			sheet.conditionalFormats.push(makeCf('A1:A4', makeRule('aboveAverage', { style: greenFill })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeUndefined()
			expect(result.get('A4')).toBeUndefined()
		})
	})

	// ── duplicateValues ──────────────────────────────────────────────

	describe('duplicateValues', () => {
		test('highlights cells whose value appears more than once', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(1))
			setCell(sheet, 1, 0, numberValue(2))
			setCell(sheet, 2, 0, numberValue(1))
			setCell(sheet, 3, 0, numberValue(3))

			sheet.conditionalFormats.push(
				makeCf('A1:A4', makeRule('duplicateValues', { style: redFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A3')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
			expect(result.get('A4')).toBeUndefined()
		})

		test('empty cells are not treated as duplicates', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, EMPTY)
			setCell(sheet, 1, 0, EMPTY)
			setCell(sheet, 2, 0, numberValue(5))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('duplicateValues', { style: redFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(0)
		})

		test('text duplicates are matched case-insensitively', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('Apple'))
			setCell(sheet, 1, 0, stringValue('apple'))
			setCell(sheet, 2, 0, stringValue('Pear'))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('duplicateValues', { style: redFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('text cells with wildcard criteria follow Excel duplicate matching', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('AL*'))
			setCell(sheet, 1, 0, stringValue('Alpha'))
			setCell(sheet, 2, 0, stringValue('Beta'))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('duplicateValues', { style: redFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
			expect(result.get('A3')).toBeUndefined()
		})
	})

	// ── uniqueValues ─────────────────────────────────────────────────

	describe('uniqueValues', () => {
		test('highlights cells whose value appears exactly once', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(1))
			setCell(sheet, 1, 0, numberValue(2))
			setCell(sheet, 2, 0, numberValue(1))
			setCell(sheet, 3, 0, numberValue(3))

			sheet.conditionalFormats.push(makeCf('A1:A4', makeRule('uniqueValues', { style: greenFill })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A4')).toBeDefined()
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('empty cells are not treated as unique', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, EMPTY)
			setCell(sheet, 1, 0, numberValue(5))

			sheet.conditionalFormats.push(makeCf('A1:A2', makeRule('uniqueValues', { style: greenFill })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A1')).toBeUndefined()
		})

		test('case-only text variants are not unique', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('North'))
			setCell(sheet, 1, 0, stringValue('north'))
			setCell(sheet, 2, 0, stringValue('South'))

			sheet.conditionalFormats.push(makeCf('A1:A3', makeRule('uniqueValues', { style: greenFill })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(1)
			expect(result.get('A3')).toBeDefined()
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeUndefined()
		})

		test('text cells with wildcard criteria follow Excel unique matching', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('AL*'))
			setCell(sheet, 1, 0, stringValue('Alpha'))
			setCell(sheet, 2, 0, stringValue('Beta'))

			sheet.conditionalFormats.push(makeCf('A1:A3', makeRule('uniqueValues', { style: greenFill })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeDefined()
		})
	})

	// ── containsText ─────────────────────────────────────────────────

	describe('containsText', () => {
		test('matches substring anywhere in cell', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('hello world'))
			setCell(sheet, 1, 0, stringValue('foo bar'))
			setCell(sheet, 2, 0, stringValue('hello'))
			setCell(sheet, 3, 0, stringValue('goodbye'))

			sheet.conditionalFormats.push(
				makeCf('A1:A4', makeRule('containsText', { formulas: ['"hello"'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
			expect(result.get('A3')).toBeDefined()
			expect(result.get('A4')).toBeUndefined()
		})

		test('no formula matches nothing', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('anything'))

			sheet.conditionalFormats.push(makeCf('A1', makeRule('containsText', { formulas: [] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
		})

		test('matches numeric cell converted to string', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(12345))

			sheet.conditionalFormats.push(makeCf('A1', makeRule('containsText', { formulas: ['"234"'] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
		})

		test('keeps numeric-looking formulas as literal text patterns', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))
			setCell(sheet, 1, 0, numberValue(23))

			sheet.conditionalFormats.push(makeCf('A1:A2', makeRule('containsText', { formulas: ['0'] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})

		test('matches OOXML boolean search formula relative to each cell', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('Dairy'))
			setCell(sheet, 1, 0, stringValue('Grain'))
			setCell(sheet, 2, 0, stringValue('Produce'))

			sheet.conditionalFormats.push(
				makeCf(
					'A1:A3',
					makeRule('containsText', {
						formulas: ['NOT(ISERROR(SEARCH("Grain",A1)))'],
					}),
				),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('matches cfRule text attribute without formula payload case-insensitively', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('North Grain'))
			setCell(sheet, 1, 0, stringValue('produce'))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('containsText', { formulas: [], text: 'grain' })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})
	})

	// ── notContainsText ──────────────────────────────────────────────

	describe('notContainsText', () => {
		test('matches cells that do not contain the pattern', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('apple'))
			setCell(sheet, 1, 0, stringValue('banana'))
			setCell(sheet, 2, 0, stringValue('pineapple'))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('notContainsText', { formulas: ['"apple"'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('no formula matches everything', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('anything'))

			sheet.conditionalFormats.push(makeCf('A1', makeRule('notContainsText', { formulas: [] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
		})

		test('uses cfRule text attribute as fallback pattern', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('Apple'))
			setCell(sheet, 1, 0, stringValue('pear'))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('notContainsText', { formulas: [], text: 'apple' })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
		})
	})

	// ── beginsWith ───────────────────────────────────────────────────

	describe('beginsWith', () => {
		test('matches cells starting with prefix', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('hello'))
			setCell(sheet, 1, 0, stringValue('world'))
			setCell(sheet, 2, 0, stringValue('hello world'))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('beginsWith', { formulas: ['"hel"'] })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
			expect(result.get('A3')).toBeDefined()
		})

		test('no formula matches nothing', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('hello'))

			sheet.conditionalFormats.push(makeCf('A1', makeRule('beginsWith', { formulas: [] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
		})

		test('uses cfRule text attribute as case-insensitive prefix', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('Hello'))
			setCell(sheet, 1, 0, stringValue('world'))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('beginsWith', { formulas: [], text: 'hel' })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})
	})

	// ── endsWith ─────────────────────────────────────────────────────

	describe('endsWith', () => {
		test('matches cells ending with suffix', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('hello'))
			setCell(sheet, 1, 0, stringValue('world'))
			setCell(sheet, 2, 0, stringValue('hello world'))

			sheet.conditionalFormats.push(makeCf('A1:A3', makeRule('endsWith', { formulas: ['"orld"'] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeDefined()
		})

		test('no formula matches nothing', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('hello'))

			sheet.conditionalFormats.push(makeCf('A1', makeRule('endsWith', { formulas: [] })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
		})

		test('uses cfRule text attribute as case-insensitive suffix', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('HELLO'))
			setCell(sheet, 1, 0, stringValue('world'))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('endsWith', { formulas: [], text: 'lo' })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})
	})

	// ── containsBlanks ───────────────────────────────────────────────

	describe('containsBlanks', () => {
		test('matches empty and empty-string cells', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(1))
			setCell(sheet, 1, 0, EMPTY)
			setCell(sheet, 2, 0, stringValue(''))
			setCell(sheet, 3, 0, stringValue('x'))

			sheet.conditionalFormats.push(makeCf('A1:A4', makeRule('containsBlanks')))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A2')?.[0]).toMatchObject({ type: 'containsBlanks' })
			expect(result.get('A3')).toBeDefined()
			expect(result.get('A4')).toBeUndefined()
		})

		test('boolean and error values are not blank', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, booleanValue(false))
			setCell(sheet, 1, 0, errorValue('#N/A'))

			sheet.conditionalFormats.push(makeCf('A1:A2', makeRule('containsBlanks')))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeUndefined()
		})
	})

	// ── notContainsBlanks ────────────────────────────────────────────

	describe('notContainsBlanks', () => {
		test('matches non-empty cells', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, EMPTY)
			setCell(sheet, 1, 0, numberValue(42))
			setCell(sheet, 2, 0, stringValue('hello'))
			setCell(sheet, 3, 0, stringValue(''))

			sheet.conditionalFormats.push(makeCf('A1:A4', makeRule('notContainsBlanks')))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeDefined()
			expect(result.get('A4')).toBeUndefined() // empty string is blank
		})
	})

	// ── timePeriod ───────────────────────────────────────────────────

	describe('timePeriod', () => {
		const EXCEL_EPOCH_OFFSET = 25569
		const MS_PER_DAY = 86400000

		function currentSerial(): number {
			const now = new Date()
			const utcMidnight = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())
			return Math.floor(utcMidnight / MS_PER_DAY) + EXCEL_EPOCH_OFFSET
		}

		test('today matches cell with today serial', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const today = currentSerial()
			setCell(sheet, 0, 0, numberValue(today))
			setCell(sheet, 1, 0, numberValue(today - 1))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('timePeriod', { timePeriod: 'today' })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})

		test('yesterday matches cell with yesterday serial', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const today = currentSerial()
			setCell(sheet, 0, 0, numberValue(today))
			setCell(sheet, 1, 0, numberValue(today - 1))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('timePeriod', { timePeriod: 'yesterday' })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
		})

		test('tomorrow matches cell with tomorrow serial', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const today = currentSerial()
			setCell(sheet, 0, 0, numberValue(today + 1))
			setCell(sheet, 1, 0, numberValue(today))

			sheet.conditionalFormats.push(
				makeCf('A1:A2', makeRule('timePeriod', { timePeriod: 'tomorrow' })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
		})

		test('last7Days matches cells within past 7 days', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const today = currentSerial()
			setCell(sheet, 0, 0, numberValue(today))
			setCell(sheet, 1, 0, numberValue(today - 6))
			setCell(sheet, 2, 0, numberValue(today - 7))
			setCell(sheet, 3, 0, numberValue(today + 1))

			sheet.conditionalFormats.push(
				makeCf('A1:A4', makeRule('timePeriod', { timePeriod: 'last7Days' })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeUndefined()
			expect(result.get('A4')).toBeUndefined()
		})

		test('thisMonth matches cells in the same month', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const today = currentSerial()
			const now = new Date()
			const firstOfMonth =
				Math.floor(Date.UTC(now.getFullYear(), now.getMonth(), 1) / MS_PER_DAY) + EXCEL_EPOCH_OFFSET
			const prevMonthDay = firstOfMonth - 1

			setCell(sheet, 0, 0, numberValue(today))
			setCell(sheet, 1, 0, numberValue(firstOfMonth))
			setCell(sheet, 2, 0, numberValue(prevMonthDay))

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('timePeriod', { timePeriod: 'thisMonth' })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeUndefined()
		})

		test('non-numeric cells do not match timePeriod', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, stringValue('hello'))

			sheet.conditionalFormats.push(makeCf('A1', makeRule('timePeriod', { timePeriod: 'today' })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(0)
		})

		test('timePeriod falls back to operator field', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			const today = currentSerial()
			setCell(sheet, 0, 0, numberValue(today))

			sheet.conditionalFormats.push(makeCf('A1', makeRule('timePeriod', { operator: 'today' })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.get('A1')).toBeDefined()
		})
	})

	// ── containsErrors ───────────────────────────────────────────────

	describe('containsErrors', () => {
		test('matches cells with error values', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, errorValue('#DIV/0!'))
			setCell(sheet, 1, 0, numberValue(42))
			setCell(sheet, 2, 0, errorValue('#N/A'))

			sheet.conditionalFormats.push(makeCf('A1:A3', makeRule('containsErrors', { style: redFill })))

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
			expect(result.get('A1')).toBeDefined()
			expect(result.get('A2')).toBeUndefined()
			expect(result.get('A3')).toBeDefined()
		})
	})

	// ── notContainsErrors ────────────────────────────────────────────

	describe('notContainsErrors', () => {
		test('matches cells without error values', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, errorValue('#DIV/0!'))
			setCell(sheet, 1, 0, numberValue(42))
			setCell(sheet, 2, 0, EMPTY)

			sheet.conditionalFormats.push(
				makeCf('A1:A3', makeRule('notContainsErrors', { style: greenFill })),
			)

			const result = evaluateConditionalFormats(sheet, wb)
			expect(result.size).toBe(2)
			expect(result.get('A1')).toBeUndefined()
			expect(result.get('A2')).toBeDefined()
			expect(result.get('A3')).toBeDefined()
		})
	})

	// ── stopIfTrue and priority interaction ──────────────────────────

	describe('stopIfTrue and priority', () => {
		test('stopIfTrue prevents later rules from matching', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))

			sheet.conditionalFormats.push({
				sqref: 'A1',
				rules: [
					{
						type: 'cellIs',
						operator: 'greaterThan',
						formulas: ['5'],
						priority: 1,
						style: greenFill,
						stopIfTrue: true,
					},
					{
						type: 'cellIs',
						operator: 'greaterThan',
						formulas: ['3'],
						priority: 2,
						style: redFill,
					},
				],
			})

			const result = evaluateConditionalFormats(sheet, wb)
			const matches = result.get('A1')
			expect(matches).toBeDefined()
			expect(matches).toHaveLength(1)
			expect(matches?.[0]?.format).toEqual(greenFill)
		})

		test('rules evaluated in priority order', () => {
			const wb = createWorkbook()
			const sheet = wb.addSheet('Sheet1')
			setCell(sheet, 0, 0, numberValue(10))

			sheet.conditionalFormats.push({
				sqref: 'A1',
				rules: [
					{
						type: 'cellIs',
						operator: 'greaterThan',
						formulas: ['5'],
						priority: 2,
						style: redFill,
					},
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
			const matches = result.get('A1')
			expect(matches).toBeDefined()
			expect(matches).toHaveLength(2)
			expect(matches?.[0]?.priority).toBe(1)
			expect(matches?.[1]?.priority).toBe(2)
		})
	})

	// ── cells outside sqref ──────────────────────────────────────────

	test('cells outside sqref are not evaluated', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, numberValue(100))
		setCell(sheet, 0, 1, numberValue(100))

		sheet.conditionalFormats.push(
			makeCf('A1', makeRule('cellIs', { operator: 'greaterThan', formulas: ['50'] })),
		)

		const result = evaluateConditionalFormats(sheet, wb)
		expect(result.get('A1')).toBeDefined()
		expect(result.get('B1')).toBeUndefined()
	})

	// ── multiple rules on same range ─────────────────────────────────

	test('multiple rules can match the same cell', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, numberValue(10))

		sheet.conditionalFormats.push({
			sqref: 'A1',
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['5'],
					priority: 1,
					style: greenFill,
				},
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['3'],
					priority: 2,
					style: boldStyle,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)
		const matches = result.get('A1')
		expect(matches).toHaveLength(2)
	})

	// ── disjoint sqref ranges ────────────────────────────────────────

	test('disjoint sqref ranges (space-separated)', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, numberValue(10))
		setCell(sheet, 0, 1, numberValue(1))
		setCell(sheet, 0, 2, numberValue(10))

		sheet.conditionalFormats.push({
			sqref: 'A1 C1',
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['5'],
					priority: 1,
					style: greenFill,
				},
			],
		})

		const result = evaluateConditionalFormats(sheet, wb)
		expect(result.get('A1')).toBeDefined()
		expect(result.get('B1')).toBeUndefined()
		expect(result.get('C1')).toBeDefined()
	})

	// ── empty conditional formats array ──────────────────────────────

	test('no conditional formats returns empty map', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		setCell(sheet, 0, 0, numberValue(42))

		const result = evaluateConditionalFormats(sheet, wb)
		expect(result.size).toBe(0)
	})
})
