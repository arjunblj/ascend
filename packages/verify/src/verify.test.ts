import { describe, expect, test } from 'bun:test'
import { createTableId, createWorkbook, type StyleId, type Table } from '@ascend/core'
import { EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'
import { check } from './checker.ts'
import { lint } from './linter.ts'
import { trace } from './tracer.ts'

const SID = 0 as StyleId

function makeCleanWorkbook() {
	const wb = createWorkbook()
	const s = wb.addSheet('Sheet1')
	s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: SID })
	s.cells.set(0, 1, { value: numberValue(2), formula: null, styleId: SID })
	s.cells.set(0, 2, { value: numberValue(3), formula: 'A1+B1', styleId: SID })
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
		s.cells.set(0, 0, { value: errorValue('#DIV/0!'), formula: '1/0', styleId: SID })
		const result = check(wb)
		expect(result.passed).toBe(false)
		const errorIssues = result.issues.filter((i) => i.rule === 'formula-errors')
		expect(errorIssues.length).toBeGreaterThanOrEqual(1)
		expect(errorIssues[0]?.message).toContain('#DIV/0!')
	})

	test('detects circular refs', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: EMPTY, formula: 'B1', styleId: SID })
		s.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: SID })
		const result = check(wb)
		expect(result.passed).toBe(false)
		const circIssues = result.issues.filter((i) => i.rule === 'circular-refs')
		expect(circIssues.length).toBeGreaterThanOrEqual(1)
	})

	test('detects broken refs to non-existent sheets', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: EMPTY, formula: 'MissingSheet!A1', styleId: SID })
		const result = check(wb)
		expect(result.passed).toBe(false)
		const brokenIssues = result.issues.filter((i) => i.rule === 'broken-refs')
		expect(brokenIssues.length).toBeGreaterThanOrEqual(1)
		expect(brokenIssues[0]?.message).toContain('MissingSheet')
	})

	test('detects invalid 3D sheet spans', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		wb.addSheet('Sheet2')
		s1.cells.set(0, 0, { value: EMPTY, formula: 'SUM(Sheet2:Sheet1!A1)', styleId: SID })
		const result = check(wb)
		expect(result.passed).toBe(false)
		const brokenIssues = result.issues.filter((issue) => issue.rule === 'broken-refs')
		expect(brokenIssues.some((issue) => issue.message.includes('Invalid 3D sheet span'))).toBe(true)
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

	test('detects external workbook references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: EMPTY, formula: '[Book1.xlsx]Sheet1!A1', styleId: SID })
		const result = check(wb)
		const externalIssues = result.issues.filter((i) => i.rule === 'external-refs')
		expect(externalIssues.length).toBeGreaterThanOrEqual(1)
		expect(externalIssues[0]?.message).toContain('External workbook reference')
		expect(externalIssues[0]?.suggestedFix).toBeDefined()
	})

	test('suggests closest sheet name for broken refs', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.addSheet('Summary')
		const s = wb.sheets[0] as (typeof wb.sheets)[0]
		s.cells.set(0, 0, { value: EMPTY, formula: 'Sumary!A1', styleId: SID })
		const result = check(wb)
		const brokenIssues = result.issues.filter((i) => i.rule === 'broken-refs')
		expect(brokenIssues.length).toBeGreaterThanOrEqual(1)
		expect(brokenIssues[0]?.suggestedFix).toContain('Summary')
	})

	test('includes suggestedFix for #REF! errors', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: errorValue('#REF!'), formula: 'A2', styleId: SID })
		const result = check(wb)
		const refErrors = result.issues.filter(
			(i) => i.rule === 'formula-errors' && i.message.includes('#REF!'),
		)
		expect(refErrors.length).toBeGreaterThanOrEqual(1)
		expect(refErrors[0]?.suggestedFix).toContain('referenced cells')
	})

	test('includes suggestedFix for circular refs', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: EMPTY, formula: 'B1', styleId: SID })
		s.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: SID })
		const result = check(wb)
		const circIssues = result.issues.filter((i) => i.rule === 'circular-refs')
		expect(circIssues.length).toBeGreaterThanOrEqual(1)
		expect(circIssues[0]?.suggestedFix).toContain('Break the cycle')
	})

	test('includes suggestedFix for #DIV/0! errors', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: errorValue('#DIV/0!'), formula: '1/0', styleId: SID })
		const result = check(wb)
		const divErrors = result.issues.filter(
			(i) => i.rule === 'formula-errors' && i.message.includes('#DIV/0!'),
		)
		expect(divErrors.length).toBeGreaterThanOrEqual(1)
		expect(divErrors[0]?.suggestedFix).toContain('check for zero')
	})

	test('includes suggestedFix for orphaned names', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.addSheet('DataSheet')
		wb.definedNames.set('MyRange', 'DtaSheet!A1:B5')
		const result = check(wb)
		const orphanIssues = result.issues.filter((i) => i.rule === 'orphaned-names')
		expect(orphanIssues.length).toBeGreaterThanOrEqual(1)
		expect(orphanIssues[0]?.suggestedFix).toContain('DataSheet')
	})
})

describe('linter', () => {
	test('detects volatile overuse', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		for (let i = 0; i < 12; i++) {
			s.cells.set(i, 0, { value: numberValue(0), formula: 'NOW()', styleId: SID })
		}
		const result = lint(wb)
		const volatileViolations = result.violations.filter((v) => v.rule === 'volatile-overuse')
		expect(volatileViolations.length).toBeGreaterThanOrEqual(1)
		expect(volatileViolations[0]?.message).toContain('12')
	})

	test('detects magic numbers', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(0), formula: 'A2*3.14159', styleId: SID })
		const result = lint(wb)
		const magic = result.violations.filter((v) => v.rule === 'hardcoded-in-formula')
		expect(magic.length).toBeGreaterThanOrEqual(1)
		expect(magic[0]?.message).toContain('3.14159')
	})

	test('does not flag 0 or 1 as magic numbers', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(0), formula: 'A2+1', styleId: SID })
		s.cells.set(0, 1, { value: numberValue(0), formula: 'A2*0', styleId: SID })
		const result = lint(wb)
		const magic = result.violations.filter((v) => v.rule === 'hardcoded-in-formula')
		expect(magic).toHaveLength(0)
	})

	test('detects fragile refs', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(0), formula: 'SUM(A1:A200)', styleId: SID })
		const result = lint(wb)
		const fragile = result.violations.filter((v) => v.rule === 'fragile-refs')
		expect(fragile.length).toBeGreaterThanOrEqual(1)
	})

	test('detects unused defined names', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: SID })
		s.cells.set(0, 1, { value: numberValue(2), formula: 'A1+1', styleId: SID })
		wb.definedNames.set('TaxRate', 'Sheet1!A1')
		wb.definedNames.set('Unused', 'Sheet1!B1')
		const result = lint(wb)
		const unused = result.violations.filter((v) => v.rule === 'unused-name')
		expect(unused.length).toBe(2)
		expect(unused.some((v) => v.message.includes('TaxRate'))).toBe(true)
		expect(unused.some((v) => v.message.includes('Unused'))).toBe(true)
	})

	test('does not flag defined names that are referenced', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(0.08), formula: null, styleId: SID })
		s.cells.set(0, 1, { value: numberValue(0), formula: 'A1*TaxRate', styleId: SID })
		wb.definedNames.set('TaxRate', 'Sheet1!A1')
		const result = lint(wb)
		const unused = result.violations.filter((v) => v.rule === 'unused-name')
		expect(unused).toHaveLength(0)
	})

	test('detects complex formulas (warning)', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, {
			value: numberValue(0),
			formula:
				'IF(A2,IF(A3,IF(A4,IF(A5,IF(A6,IF(A7,IF(A8,IF(A9,IF(A10,IF(A11,1,0),0),0),0),0),0),0),0),0),0)',
			styleId: SID,
		})
		const result = lint(wb)
		const complex = result.violations.filter((v) => v.rule === 'complex-formula')
		expect(complex.length).toBeGreaterThanOrEqual(1)
		expect(complex[0]?.severity).toBe('warning')
	})

	test('detects complex formulas (error at depth > 20)', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		let formula = '1'
		for (let i = 0; i < 21; i++) {
			formula = `IF(TRUE,${formula},0)`
		}
		s.cells.set(0, 0, { value: numberValue(0), formula, styleId: SID })
		const result = lint(wb)
		const complex = result.violations.filter((v) => v.rule === 'complex-formula')
		expect(complex.length).toBeGreaterThanOrEqual(1)
		expect(complex[0]?.severity).toBe('error')
	})

	test('does not flag shallow formulas as complex', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(0), formula: 'SUM(A2:A10)', styleId: SID })
		const result = lint(wb)
		const complex = result.violations.filter((v) => v.rule === 'complex-formula')
		expect(complex).toHaveLength(0)
	})

	test('respects custom volatile threshold', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		for (let i = 0; i < 4; i++) {
			s.cells.set(i, 0, { value: numberValue(0), formula: 'NOW()', styleId: SID })
		}
		const defaultResult = lint(wb)
		expect(defaultResult.violations.filter((v) => v.rule === 'volatile-overuse')).toHaveLength(0)

		const customResult = lint(wb, undefined, { volatileThreshold: 3 })
		const volatile = customResult.violations.filter((v) => v.rule === 'volatile-overuse')
		expect(volatile.length).toBeGreaterThanOrEqual(1)
	})

	test('respects custom complexity thresholds', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, {
			value: numberValue(0),
			formula: 'IF(A2,IF(A3,IF(A4,IF(A5,1,0),0),0),0)',
			styleId: SID,
		})
		const defaultResult = lint(wb)
		expect(defaultResult.violations.filter((v) => v.rule === 'complex-formula')).toHaveLength(0)

		const customResult = lint(wb, undefined, { complexityDepthWarning: 3 })
		const complex = customResult.violations.filter((v) => v.rule === 'complex-formula')
		expect(complex.length).toBeGreaterThanOrEqual(1)
	})

	test('respects custom fragile ref threshold', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(0), formula: 'SUM(A1:A20)', styleId: SID })

		const defaultResult = lint(wb)
		expect(defaultResult.violations.filter((v) => v.rule === 'fragile-refs')).toHaveLength(0)

		const customResult = lint(wb, undefined, { fragileRefThreshold: 10 })
		const fragile = customResult.violations.filter((v) => v.rule === 'fragile-refs')
		expect(fragile.length).toBeGreaterThanOrEqual(1)
	})
})

describe('tracer', () => {
	test('returns precedents and dependents', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(10), formula: null, styleId: SID })
		s.cells.set(0, 1, { value: numberValue(20), formula: null, styleId: SID })
		s.cells.set(0, 2, { value: numberValue(30), formula: 'A1+B1', styleId: SID })

		const result = trace(wb, 'Sheet1', 'C1')
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error(result.error.message)

		expect(result.value.formula).toBe('A1+B1')
		expect(result.value.precedents.length).toBe(2)
		expect(result.value.dependents).toHaveLength(0)
	})

	test('returns empty for non-formula cell', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: stringValue('hello'), formula: null, styleId: SID })

		const result = trace(wb, 'Sheet1', 'A1')
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error(result.error.message)

		expect(result.value.formula).toBeNull()
		expect(result.value.precedents).toHaveLength(0)
	})

	test('returns dependents of a source cell', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(5), formula: null, styleId: SID })
		s.cells.set(0, 1, { value: numberValue(10), formula: 'A1*2', styleId: SID })
		s.cells.set(0, 2, { value: numberValue(15), formula: 'A1+B1', styleId: SID })

		const result = trace(wb, 'Sheet1', 'A1')
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error(result.error.message)

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
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: SID })
		s.cells.set(0, 1, { value: numberValue(1), formula: 'A1', styleId: SID })
		s.cells.set(0, 2, { value: numberValue(1), formula: 'B1', styleId: SID })
		s.cells.set(0, 3, { value: numberValue(1), formula: 'C1', styleId: SID })

		const result = trace(wb, 'Sheet1', 'D1', { maxDepth: 1 })
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error(result.error.message)

		expect(result.value.precedents.length).toBe(1)
		expect(result.value.precedents[0]?.ref).toBe('C1')
	})

	test('reports whole-column precedents symbolically instead of expanding them', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: SID })
		s.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: SID })
		s.cells.set(0, 1, { value: numberValue(3), formula: 'SUM(A:A)', styleId: SID })

		const result = trace(wb, 'Sheet1', 'B1')
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error(result.error.message)

		expect(result.value.precedents).toHaveLength(1)
		expect(result.value.precedents[0]?.ref).toBe('A:A')
	})

	test('returns null cyclePath for non-circular cell', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: SID })
		s.cells.set(0, 1, { value: numberValue(2), formula: 'A1+1', styleId: SID })

		const result = trace(wb, 'Sheet1', 'B1')
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error(result.error.message)

		expect(result.value.cyclePath).toBeNull()
	})

	test('returns cyclePath for circular reference', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: EMPTY, formula: 'B1', styleId: SID })
		s.cells.set(0, 1, { value: EMPTY, formula: 'A1', styleId: SID })

		const result = trace(wb, 'Sheet1', 'A1')
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error(result.error.message)

		expect(result.value.cyclePath).not.toBeNull()
		const path = result.value.cyclePath ?? []
		expect(path.length).toBeGreaterThanOrEqual(2)
		expect(path.some((r) => r.includes('A1'))).toBe(true)
		expect(path.some((r) => r.includes('B1'))).toBe(true)
	})

	test('reports structured-reference precedents from resolved dependency metadata', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: stringValue('Player'), formula: null, styleId: SID })
		s.cells.set(0, 1, { value: stringValue('Score'), formula: null, styleId: SID })
		s.cells.set(1, 0, { value: stringValue('Mina'), formula: null, styleId: SID })
		s.cells.set(1, 1, { value: numberValue(10), formula: null, styleId: SID })
		s.cells.set(2, 0, { value: stringValue('Noah'), formula: null, styleId: SID })
		s.cells.set(2, 1, { value: numberValue(12), formula: null, styleId: SID })
		s.tables.push({
			id: createTableId(),
			name: 'Scores',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Player' }, { name: 'Score' }],
			hasHeaders: true,
			hasTotals: false,
		})
		s.cells.set(4, 0, { value: EMPTY, formula: 'SUM(Scores[Score])', styleId: SID })

		const result = trace(wb, 'Sheet1', 'A5')
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error(result.error.message)

		expect(result.value.precedents).toHaveLength(1)
		expect(result.value.precedents[0]?.ref).toBe('B2:B3')
	})
})
