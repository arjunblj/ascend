import { describe, expect, test } from 'bun:test'
import { createTableId, createWorkbook, type Table } from '@ascend/core'
import { EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import { check } from './checker.ts'
import { lint } from './linter.ts'
import { trace } from './tracer.ts'

function makeCleanWorkbook() {
	const wb = createWorkbook()
	const s = wb.addSheet('Sheet1')
	s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: 0 })
	s.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: 0 })
	s.cells.set(0, 2, { value: numberValue(3), formula: 'A1+B1', styleId: 0 })
	return wb
}

describe('checker', () => {
	test('passes on clean workbook', () => {
		const wb = makeCleanWorkbook()
		const result = check(wb)
		expect(result.passed).toBe(true)
		expect(result.issues).toHaveLength(0)
	})

	test('detects formula errors', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: errorValue('#DIV/0!'), formula: '1/0', styleId: 0 })
		const result = check(wb)
		expect(result.passed).toBe(false)
		const errorIssues = result.issues.filter((i) => i.rule === 'formula-errors')
		expect(errorIssues.length).toBeGreaterThanOrEqual(1)
		expect(errorIssues[0]?.message).toContain('#DIV/0!')
	})

	test('detects circular refs', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: EMPTY, formula: 'B1', styleId: 0 })
		s.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: 0 })
		const result = check(wb)
		expect(result.passed).toBe(false)
		const circIssues = result.issues.filter((i) => i.rule === 'circular-refs')
		expect(circIssues.length).toBeGreaterThanOrEqual(1)
	})

	test('detects broken refs to non-existent sheets', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: EMPTY, formula: 'MissingSheet!A1', styleId: 0 })
		const result = check(wb)
		expect(result.passed).toBe(false)
		const brokenIssues = result.issues.filter((i) => i.rule === 'broken-refs')
		expect(brokenIssues.length).toBeGreaterThanOrEqual(1)
		expect(brokenIssues[0]?.message).toContain('MissingSheet')
	})

	test('detects orphaned names', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.definedNames.set('MyRange', 'DeletedSheet!A1:B5')
		const result = check(wb)
		expect(result.passed).toBe(false)
		const orphanIssues = result.issues.filter((i) => i.rule === 'orphaned-names')
		expect(orphanIssues.length).toBeGreaterThanOrEqual(1)
	})

	test('detects table integrity issues', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		const table: Table = {
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 5, col: 3 } },
			columns: [{ name: 'A' }, { name: 'B' }],
			hasHeaders: true,
			hasTotals: false,
		}
		s.tables.push(table)
		const result = check(wb)
		expect(result.passed).toBe(false)
		const tableIssues = result.issues.filter((i) => i.rule === 'table-integrity')
		expect(tableIssues.length).toBeGreaterThanOrEqual(1)
	})
})

describe('linter', () => {
	test('detects volatile overuse', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		for (let i = 0; i < 12; i++) {
			s.cells.set(i, 0, { value: numberValue(0), formula: 'NOW()', styleId: 0 })
		}
		const result = lint(wb)
		const volatileViolations = result.violations.filter((v) => v.rule === 'volatile-overuse')
		expect(volatileViolations.length).toBeGreaterThanOrEqual(1)
		expect(volatileViolations[0]?.message).toContain('12')
	})

	test('detects magic numbers', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(0), formula: 'A2*3.14159', styleId: 0 })
		const result = lint(wb)
		const magic = result.violations.filter((v) => v.rule === 'hardcoded-in-formula')
		expect(magic.length).toBeGreaterThanOrEqual(1)
		expect(magic[0]?.message).toContain('3.14159')
	})

	test('does not flag 0 or 1 as magic numbers', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(0), formula: 'A2+1', styleId: 0 })
		s.cells.set(0, 1, { value: numberValue(0), formula: 'A2*0', styleId: 0 })
		const result = lint(wb)
		const magic = result.violations.filter((v) => v.rule === 'hardcoded-in-formula')
		expect(magic).toHaveLength(0)
	})

	test('detects fragile refs', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(0), formula: 'SUM(A1:A200)', styleId: 0 })
		const result = lint(wb)
		const fragile = result.violations.filter((v) => v.rule === 'fragile-refs')
		expect(fragile.length).toBeGreaterThanOrEqual(1)
	})
})

describe('tracer', () => {
	test('returns precedents and dependents', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: 0 })
		s.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: 0 })
		s.cells.set(0, 2, { value: numberValue(30), formula: 'A1+B1', styleId: 0 })

		const result = trace(wb, 'Sheet1', 'C1')
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.formula).toBe('A1+B1')
		expect(result.value.precedents.length).toBe(2)
		expect(result.value.dependents).toHaveLength(0)
	})

	test('returns empty for non-formula cell', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: stringValue('hello'), formula: null, styleId: 0 })

		const result = trace(wb, 'Sheet1', 'A1')
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.formula).toBeNull()
		expect(result.value.precedents).toHaveLength(0)
	})

	test('returns dependents of a source cell', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: 0 })
		s.cells.set(0, 1, { value: numberValue(10), formula: 'A1*2', styleId: 0 })
		s.cells.set(0, 2, { value: numberValue(15), formula: 'A1+B1', styleId: 0 })

		const result = trace(wb, 'Sheet1', 'A1')
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.dependents.length).toBe(2)
	})

	test('returns error for invalid sheet', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		const result = trace(wb, 'NoSuchSheet', 'A1')
		expect(result.ok).toBe(false)
	})

	test('respects maxDepth', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: 0 })
		s.cells.set(0, 1, { value: numberValue(1), formula: 'A1', styleId: 0 })
		s.cells.set(0, 2, { value: numberValue(1), formula: 'B1', styleId: 0 })
		s.cells.set(0, 3, { value: numberValue(1), formula: 'C1', styleId: 0 })

		const result = trace(wb, 'Sheet1', 'D1', { maxDepth: 1 })
		expect(result.ok).toBe(true)
		if (!result.ok) return

		expect(result.value.precedents.length).toBe(1)
		expect(result.value.precedents[0]?.ref).toBe('C1')
	})
})
