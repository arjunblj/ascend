import { describe, expect, test } from 'bun:test'
import {
	createSheetId,
	createTableId,
	createWorkbook,
	type StyleId,
	type Table,
} from '@ascend/core'
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

	test('reports machine-readable blocked spill diagnostics', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, {
			value: errorValue('#SPILL!'),
			formula: 'SEQUENCE(3)',
			styleId: SID,
			formulaInfo: {
				kind: 'blockedSpill',
				anchorRef: 'Sheet1!A1',
				ref: 'A1:A3',
				blockingRefs: ['A2'],
			},
		})
		s.cells.set(1, 0, { value: stringValue('blocker'), formula: null, styleId: SID })

		const result = check(wb)
		const spillIssue = result.issues.find((issue) => issue.rule === 'spill-diagnostics')
		expect(result.passed).toBe(false)
		expect(spillIssue?.refs).toEqual(['Sheet1!A1', 'Sheet1!A2'])
		expect(spillIssue?.details).toEqual({
			error: '#SPILL!',
			cause: 'occupied-cell',
			spillRange: 'Sheet1!A1:A3',
			blockingRefs: ['Sheet1!A2'],
		})
	})

	test('refreshes blocked spill diagnostics for stale imported caches', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, {
			value: errorValue('#SPILL!'),
			formula: 'SEQUENCE(3)',
			styleId: SID,
		})
		s.cells.set(1, 0, { value: stringValue('blocker'), formula: null, styleId: SID })

		const result = check(wb)
		const spillIssue = result.issues.find((issue) => issue.rule === 'spill-diagnostics')
		expect(spillIssue?.details).toEqual({
			error: '#SPILL!',
			cause: 'occupied-cell',
			spillRange: 'Sheet1!A1:A3',
			blockingRefs: ['Sheet1!A2'],
		})
		expect(s.cells.get(0, 0)?.formulaInfo).toBeUndefined()
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

	test('detects ambiguous overlapping conditional format priorities', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.conditionalFormats.push(
			{
				sqref: 'A1:A5',
				rules: [{ type: 'expression', priority: 1, formulas: ['A1>0'] }],
			},
			{
				sqref: 'A3:A7',
				rules: [{ type: 'cellIs', priority: 1, formulas: ['3'] }],
			},
		)
		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'conditional-format-integrity')
		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('share priority 1')
		expect(issues[0]?.refs).toEqual(['Sheet1!A1:A5', 'Sheet1!A3:A7'])
		expect(issues[0]?.details).toMatchObject({
			priority: 1,
			left: { source: 'conditionalFormat', sqref: 'A1:A5', ruleType: 'expression' },
			right: { source: 'conditionalFormat', sqref: 'A3:A7', ruleType: 'cellIs' },
		})
	})

	test('detects overlapping legacy and x14 conditional format priority collisions', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.conditionalFormats.push({
			sqref: 'B2:B5',
			rules: [{ type: 'expression', priority: 2, formulas: ['B2>0'] }],
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B4:B8',
			priority: 2,
			type: 'dataBar',
			formulas: [],
		})
		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'conditional-format-integrity')
		expect(issues).toHaveLength(1)
		expect(issues[0]?.details).toMatchObject({
			priority: 2,
			left: { source: 'conditionalFormat', sqref: 'B2:B5' },
			right: { source: 'x14ConditionalFormat', sqref: 'B4:B8', ruleType: 'dataBar' },
		})
	})

	test('detects duplicate threaded comment ids in the same part', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push(
			{
				ref: 'A1',
				text: 'Root',
				id: '{thread-1}',
				personId: '{person-1}',
				author: 'Alex',
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
			{
				ref: 'A2',
				text: 'Unexpected duplicate',
				id: '{thread-1}',
				personId: '{person-1}',
				author: 'Alex',
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
		)

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'threaded-comment-integrity')
		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('Duplicate threaded comment id')
		expect(issues[0]?.refs).toEqual(['Sheet1!A1', 'Sheet1!A2'])
		expect(issues[0]?.details).toEqual({
			partPath: 'xl/threadedComments/threadedComment1.xml',
			id: '{thread-1}',
			firstCommentIndex: 0,
			duplicateCommentIndex: 1,
		})
	})

	test('detects threaded comment replies with missing parent ids', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push({
			ref: 'C3',
			text: 'Reply without preserved root',
			id: '{reply-1}',
			parentId: '{missing-root}',
			personId: '{person-1}',
			author: 'Alex',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'threaded-comment-integrity')
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('references missing parent id')
		expect(issues[0]?.refs).toEqual(['Sheet1!C3'])
		expect(issues[0]?.details).toMatchObject({
			partPath: 'xl/threadedComments/threadedComment1.xml',
			commentIndex: 0,
			id: '{reply-1}',
			parentId: '{missing-root}',
		})
	})

	test('detects threaded comments with unresolved person ids', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push({
			ref: 'D4',
			text: 'Author sidecar missing',
			id: '{thread-1}',
			personId: '{missing-person}',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'threaded-comment-integrity')
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('references unknown person id')
		expect(issues[0]?.details).toMatchObject({
			partPath: 'xl/threadedComments/threadedComment1.xml',
			commentIndex: 0,
			id: '{thread-1}',
			personId: '{missing-person}',
		})
	})

	test('detects legacy comment VML row and column target drift', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('B2', {
			text: 'Review',
			author: 'Ada',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 4,
				column: 3,
				anchor: [1, 15, 1, 2, 3, 15, 4, 16],
			},
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'legacy-comment-drawing-integrity')
		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('points to row 4, column 3')
		expect(issues[0]?.refs).toEqual(['Sheet1!B2'])
		expect(issues[0]?.details).toEqual({
			ref: 'B2',
			expectedRow: 1,
			expectedColumn: 1,
			actualRow: 4,
			actualColumn: 3,
			shapeId: '_x0000_s1025',
		})
	})

	test('detects duplicate legacy comment VML shape ids', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('A1', {
			text: 'One',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 0,
				column: 0,
			},
		})
		s.comments.set('C3', {
			text: 'Two',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 2,
				column: 2,
			},
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'legacy-comment-drawing-integrity')
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('Duplicate legacy comment VML shape id')
		expect(issues[0]?.refs).toEqual(['Sheet1!A1', 'Sheet1!C3'])
		expect(issues[0]?.details).toEqual({
			shapeId: '_x0000_s1025',
			firstRef: 'A1',
			duplicateRef: 'C3',
		})
	})

	test('detects legacy comment VML anchors with negative coordinates', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('A1', {
			text: 'One',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 0,
				column: 0,
				anchor: [0, 0, 0, 0, 2, 0, -1, 0],
			},
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'legacy-comment-drawing-integrity')
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('not eight non-negative integers')
		expect(issues[0]?.details).toEqual({
			ref: 'A1',
			anchor: [0, 0, 0, 0, 2, 0, -1, 0],
			shapeId: '_x0000_s1025',
		})
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

	test('detects chart series references to missing local sheets', () => {
		const wb = createWorkbook()
		wb.addSheet('Data')
		wb.addSheet('Summary')
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Summary',
			chartType: 'barChart',
			series: [
				{
					nameRef: "'Data'!$B$1",
					categoryRef: "'Missing Data'!$A$2:$A$5",
					valueRef: "'Data'!$B$2:$B$5",
				},
			],
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'chart-series-integrity')
		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('Missing Data')
		expect(issues[0]?.refs).toEqual(['xl/charts/chart1.xml#series0'])
		expect(issues[0]?.details).toEqual({
			partPath: 'xl/charts/chart1.xml',
			seriesIndex: 0,
			field: 'categoryRef',
			reference: "'Missing Data'!$A$2:$A$5",
			sheetName: 'Missing Data',
			ownerSheet: 'Summary',
			chartType: 'barChart',
		})
	})

	test('does not flag external workbook chart series references as missing local sheets', () => {
		const wb = createWorkbook()
		wb.addSheet('Data')
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Data',
			series: [
				{
					categoryRef: "'[Book1.xlsx]Sheet1'!$A$2:$A$5",
					valueRef: '[Book1.xlsx]Sheet1!$B$2:$B$5',
				},
			],
		})

		const result = check(wb)
		expect(result.issues.filter((i) => i.rule === 'chart-series-integrity')).toHaveLength(0)
	})

	test('detects chart parts without worksheet or chartsheet ownership', () => {
		const wb = createWorkbook()
		wb.addSheet('Data')
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			chartType: 'lineChart',
			series: [],
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'chart-part-ownership')
		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('not attributed')
		expect(issues[0]?.details).toEqual({
			partPath: 'xl/charts/chart1.xml',
			chartType: 'lineChart',
		})
	})

	test('accepts chartsheet-owned chart parts without worksheet ownership', () => {
		const wb = createWorkbook()
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			chartType: 'barChart',
			series: [],
		})
		wb.chartSheets.push({
			name: 'Chart 1',
			sheetId: createSheetId(),
			relId: 'rIdChartSheet',
			partPath: 'xl/chartsheets/sheet1.xml',
			state: 'visible',
			chartPartPaths: ['xl/charts/chart1.xml'],
		})

		const result = check(wb)
		expect(result.issues.filter((i) => i.rule === 'chart-part-ownership')).toHaveLength(0)
	})

	test('detects chart parts attributed to missing worksheet owners', () => {
		const wb = createWorkbook()
		wb.addSheet('Summary')
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Summry',
			series: [],
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'chart-part-ownership')
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('Summry')
		expect(issues[0]?.suggestedFix).toContain('Summary')
		expect(issues[0]?.details).toEqual({
			partPath: 'xl/charts/chart1.xml',
			ownerSheet: 'Summry',
		})
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
