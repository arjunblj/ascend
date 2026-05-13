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

function queryTableRef() {
	return {
		relationshipId: 'rIdQuery',
		partPath: 'xl/queryTables/queryTable1.xml',
		relationshipType:
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
		target: '../queryTables/queryTable1.xml',
	}
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

	test('surfaces stale calculation metadata as workbook integrity warning', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: SID })
		s.cells.set(0, 1, { value: numberValue(2), formula: 'A1*2', styleId: SID })
		wb.calcSettings = {
			...wb.calcSettings,
			calcMode: 'manual',
			fullCalcOnLoad: true,
			calcCompleted: false,
			calcOnSave: false,
			forceFullCalc: true,
		}

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/calcChain.xml',
						featureFamily: 'preservedCalcChain',
						ownerScope: 'metadata',
						preservationPolicy: 'discard-on-recalc',
					},
				],
				relationships: [],
			},
		})
		const freshnessIssue = result.issues.find((issue) => issue.rule === 'calc-freshness')
		expect(result.passed).toBe(false)
		expect(freshnessIssue).toMatchObject({
			severity: 'warning',
			refs: ['Sheet1!B1'],
			suggestedFix: expect.stringContaining('Recalculate the workbook'),
			details: {
				kind: 'stale-calculation-metadata',
				reasons: [
					'manual calculation mode',
					'full recalculation requested on load',
					'calculation not completed',
					'forced full recalculation',
				],
				formulaCount: 1,
				calcChainParts: ['xl/calcChain.xml'],
				calcMode: 'manual',
				fullCalcOnLoad: true,
				calcCompleted: false,
				calcOnSave: false,
				forceFullCalc: true,
			},
		})
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

	test('parses defined name formulas for missing local and external references', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.addSheet("O'Brien")
		wb.definedNames.set('BrokenName', 'SUM(MissingSheet!A1,Sheet1!A1)')
		wb.definedNames.set('QualifiedName', 'MissingSheet!SomeName')
		wb.definedNames.set('RepeatedQualifiedName', 'MissingSheet!SomeName+missingsheet!OtherName')
		wb.definedNames.set('TextOnly', '"MissingSheet!A1"')
		wb.definedNames.set('TextInFunction', 'IF(TRUE,"MissingSheet!SomeName","")')
		wb.definedNames.set('QuotedLocal', "'O''Brien'!A1")
		wb.definedNames.set('ExternalName', "'[Budget.xlsx]Data'!$A$1")
		wb.definedNames.set('ExternalQualifiedName', '[0]!SomeName')

		const result = check(wb)
		const orphanIssues = result.issues.filter((issue) => issue.rule === 'orphaned-names')
		const formulaIssue = orphanIssues.find((issue) => issue.details?.name === 'BrokenName')
		const qualifiedNameIssue = orphanIssues.find((issue) => issue.details?.name === 'QualifiedName')
		const repeatedQualifiedNameIssues = orphanIssues.filter(
			(issue) => issue.details?.name === 'RepeatedQualifiedName',
		)
		const externalIssue = result.issues.find(
			(issue) =>
				issue.rule === 'external-refs' && issue.details?.kind === 'defined-name-external-reference',
		)
		const externalNameIssues = result.issues.filter(
			(issue) =>
				issue.rule === 'external-refs' && issue.details?.kind === 'defined-name-external-reference',
		)

		expect(result.passed).toBe(false)
		expect(orphanIssues.map((issue) => issue.details?.name).sort()).toEqual([
			'BrokenName',
			'QualifiedName',
			'RepeatedQualifiedName',
		])
		expect(formulaIssue?.message).toContain('MissingSheet')
		expect(formulaIssue?.details).toMatchObject({
			kind: 'defined-name-missing-sheet-reference',
			name: 'BrokenName',
			missingSheet: 'MissingSheet',
		})
		expect(qualifiedNameIssue?.details).toMatchObject({
			kind: 'defined-name-missing-sheet-reference',
			name: 'QualifiedName',
			missingSheet: 'MissingSheet',
		})
		expect(repeatedQualifiedNameIssues).toHaveLength(1)
		expect(repeatedQualifiedNameIssues[0]?.details).toMatchObject({
			kind: 'defined-name-missing-sheet-reference',
			name: 'RepeatedQualifiedName',
			missingSheet: 'MissingSheet',
		})
		expect(externalIssue?.severity).toBe('warning')
		expect(externalIssue?.details).toMatchObject({
			name: 'ExternalName',
			externalTarget: '[Budget.xlsx]',
		})
		expect(externalNameIssues.map((issue) => issue.details?.name).sort()).toEqual([
			'ExternalName',
			'ExternalQualifiedName',
		])
		expect(orphanIssues.some((issue) => issue.details?.name === 'ExternalName')).toBe(false)
		expect(orphanIssues.some((issue) => issue.details?.name === 'ExternalQualifiedName')).toBe(
			false,
		)
		expect(orphanIssues.some((issue) => issue.details?.name === 'TextOnly')).toBe(false)
		expect(orphanIssues.some((issue) => issue.details?.name === 'TextInFunction')).toBe(false)
		expect(orphanIssues.some((issue) => issue.details?.name === 'QuotedLocal')).toBe(false)
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

	test('detects duplicate table names and ids across workbook table parts', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		const duplicateId = createTableId()
		s1.tables.push({
			id: duplicateId,
			name: 'Sales',
			sheetId: s1.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Region' }, { name: 'Amount' }],
			hasHeaders: true,
			hasTotals: false,
			partPath: 'xl/tables/table1.xml',
		})
		s2.tables.push({
			id: duplicateId,
			name: 'SALES',
			sheetId: s2.id,
			ref: { start: { row: 4, col: 2 }, end: { row: 6, col: 3 } },
			columns: [{ name: 'Region' }, { name: 'Amount' }],
			hasHeaders: true,
			hasTotals: false,
			partPath: 'xl/tables/table2.xml',
		})

		const result = check(wb)
		const tableIssues = result.issues.filter((i) => i.rule === 'table-integrity')

		expect(result.passed).toBe(false)
		expect(tableIssues.some((issue) => issue.details?.kind === 'duplicate-table-name')).toBe(true)
		expect(tableIssues.some((issue) => issue.details?.kind === 'duplicate-table-id')).toBe(true)
		expect(
			tableIssues.find((issue) => issue.details?.kind === 'duplicate-table-name')?.refs,
		).toEqual(['Sheet1!A1:B3', 'Sheet2!C5:D7'])
	})

	test('detects duplicate table column names case-insensitively', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'region' }],
			hasHeaders: true,
			hasTotals: false,
			partPath: 'xl/tables/table1.xml',
		})

		const result = check(wb)
		const issue = result.issues.find(
			(i) => i.rule === 'table-integrity' && i.details?.kind === 'duplicate-table-column-name',
		)

		expect(result.passed).toBe(false)
		expect(issue?.message).toContain('Sales')
		expect(issue?.refs).toEqual(['Sheet1!A1:C4', 'Sheet1!A1', 'Sheet1!C1'])
		expect(issue?.suggestedFix).toContain('Rename one of the duplicate table columns')
		expect(issue?.details).toEqual({
			kind: 'duplicate-table-column-name',
			tableName: 'Sales',
			sheetName: 'Sheet1',
			ref: 'A1:C4',
			normalizedName: 'region',
			first: { columnName: 'Region', columnIndex: 0, ref: 'A1' },
			duplicate: { columnName: 'region', columnIndex: 2, ref: 'C1' },
			partPath: 'xl/tables/table1.xml',
		})
	})

	test('detects overlapping table ranges on the same worksheet', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push(
			{
				id: createTableId(),
				name: 'Actuals',
				sheetId: s.id,
				ref: { start: { row: 0, col: 0 }, end: { row: 4, col: 2 } },
				columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Owner' }],
				hasHeaders: true,
				hasTotals: false,
				partPath: 'xl/tables/table1.xml',
			},
			{
				id: createTableId(),
				name: 'Forecast',
				sheetId: s.id,
				ref: { start: { row: 2, col: 1 }, end: { row: 6, col: 3 } },
				columns: [{ name: 'Scenario' }, { name: 'Amount' }, { name: 'Owner' }],
				hasHeaders: true,
				hasTotals: false,
				partPath: 'xl/tables/table2.xml',
			},
		)

		const result = check(wb)
		const overlapIssue = result.issues.find(
			(issue) =>
				issue.rule === 'table-integrity' && issue.details?.kind === 'overlapping-table-ranges',
		)

		expect(result.passed).toBe(false)
		expect(overlapIssue?.message).toContain('Actuals')
		expect(overlapIssue?.refs).toEqual(['Sheet1!A1:C5', 'Sheet1!B3:D7'])
		expect(overlapIssue?.suggestedFix).toContain('unambiguous table ownership')
	})

	test('detects queryTable relationship binding mismatches', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'SalesQuery',
			sheetId: s.id,
			partPath: 'xl/tables/table1.xml',
			tableType: 'queryTable',
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [
				{ name: 'Name', queryTableFieldId: 1 },
				{ name: 'Amount', queryTableFieldId: 2 },
			],
			hasHeaders: true,
			hasTotals: false,
			queryTable: queryTableRef(),
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/tables/table1.xml',
						featureFamily: 'preservedTable',
						ownerScope: 'worksheet',
					},
					{
						path: 'xl/queryTables/queryTable1.xml',
						featureFamily: 'preservedQueryTable',
						ownerScope: 'worksheet',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/tables/table1.xml',
						relationshipPartPath: 'xl/tables/_rels/table1.xml.rels',
						id: 'rIdOther',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
						rawTarget: '../queryTables/queryTable1.xml',
						resolvedTarget: 'xl/queryTables/queryTable1.xml',
						featureFamily: 'preservedQueryTable',
					},
				],
			},
		})
		const queryIssues = result.issues.filter((i) => i.rule === 'table-query-integrity')
		expect(result.passed).toBe(false)
		expect(queryIssues).toHaveLength(1)
		expect(queryIssues[0]?.message).toContain('does not bind')
		expect(queryIssues[0]?.refs).toEqual(['Sheet1!A1:B3', 'xl/queryTables/queryTable1.xml'])
		expect(queryIssues[0]?.details).toMatchObject({
			kind: 'query-table-relationship-binding-mismatch',
			table: {
				tableName: 'SalesQuery',
				tablePartPath: 'xl/tables/table1.xml',
				queryTablePartPath: 'xl/queryTables/queryTable1.xml',
				relationshipId: 'rIdQuery',
			},
			incomingRelationships: [
				{
					sourcePartPath: 'xl/tables/table1.xml',
					relationshipPartPath: 'xl/tables/_rels/table1.xml.rels',
					id: 'rIdOther',
					resolvedTarget: 'xl/queryTables/queryTable1.xml',
				},
			],
		})
	})

	test('detects missing and duplicate queryTable field bindings', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'SalesQuery',
			sheetId: s.id,
			partPath: 'xl/tables/table1.xml',
			tableType: 'queryTable',
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } },
			columns: [
				{ name: 'Name', queryTableFieldId: 1 },
				{ name: 'Amount', queryTableFieldId: 1 },
				{ name: 'Status' },
			],
			hasHeaders: true,
			hasTotals: false,
			queryTable: queryTableRef(),
		})

		const result = check(wb)
		const queryIssues = result.issues.filter((i) => i.rule === 'table-query-integrity')
		const duplicateFieldId = queryIssues.find(
			(issue) => issue.details?.kind === 'duplicate-query-table-field-id',
		)
		const missingFieldId = queryIssues.find(
			(issue) => issue.details?.kind === 'missing-query-table-field-id',
		)

		expect(result.passed).toBe(false)
		expect(duplicateFieldId?.severity).toBe('error')
		expect(duplicateFieldId?.refs).toEqual([
			'Sheet1!A1:C3',
			'Sheet1!A1',
			'Sheet1!B1',
			'xl/queryTables/queryTable1.xml',
		])
		expect(duplicateFieldId?.suggestedFix).toContain('queryTableFieldId bindings')
		expect(duplicateFieldId?.details).toMatchObject({
			kind: 'duplicate-query-table-field-id',
			queryTableFieldId: 1,
			table: {
				tableName: 'SalesQuery',
				queryTablePartPath: 'xl/queryTables/queryTable1.xml',
			},
		})
		expect(missingFieldId?.severity).toBe('error')
		expect(missingFieldId?.refs).toEqual(['Sheet1!A1:C3', 'xl/queryTables/queryTable1.xml'])
		expect(missingFieldId?.suggestedFix).toContain('Restore tableColumn queryTableFieldId')
		expect(missingFieldId?.details).toMatchObject({
			kind: 'missing-query-table-field-id',
			table: {
				tableName: 'SalesQuery',
				queryTablePartPath: 'xl/queryTables/queryTable1.xml',
			},
			columns: [{ columnName: 'Status', columnIndex: 2, ref: 'Sheet1!C1' }],
		})
	})

	test('detects orphan queryTable package sidecars', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			partPath: 'xl/tables/table1.xml',
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Amount' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/tables/table1.xml',
						featureFamily: 'preservedTable',
						ownerScope: 'worksheet',
					},
					{
						path: 'xl/queryTables/queryTable1.xml',
						featureFamily: 'preservedQueryTable',
						ownerScope: 'worksheet',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/tables/table1.xml',
						relationshipPartPath: 'xl/tables/_rels/table1.xml.rels',
						id: 'rIdQuery',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
						rawTarget: '../queryTables/queryTable1.xml',
						resolvedTarget: 'xl/queryTables/queryTable1.xml',
						featureFamily: 'preservedQueryTable',
					},
				],
			},
		})
		const orphanIssue = result.issues.find(
			(i) => i.rule === 'table-query-integrity' && i.details?.kind === 'orphan-query-table-part',
		)
		expect(result.passed).toBe(false)
		expect(orphanIssue?.severity).toBe('warning')
		expect(orphanIssue?.refs).toEqual(['xl/queryTables/queryTable1.xml'])
		expect(orphanIssue?.suggestedFix).toContain('orphan queryTable sidecar')
	})

	test('does not treat worksheet-owned queryTable connection parts as table orphans', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.connectionParts.push({
			kind: 'queryTable',
			partPath: 'xl/queryTables/queryTable1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
			relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
			sheetName: 'Sheet1',
			relationshipCount: 0,
			name: 'Query1',
			connectionId: 1,
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/queryTables/queryTable1.xml',
						featureFamily: 'preservedQueryTable',
						ownerScope: 'worksheet',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
					},
					{
						path: 'xl/queryTables/queryTable2.xml',
						featureFamily: 'preservedQueryTable',
						ownerScope: 'worksheet',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rId1',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
						rawTarget: '../queryTables/queryTable1.xml',
						resolvedTarget: 'xl/queryTables/queryTable1.xml',
						featureFamily: 'preservedQueryTable',
					},
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rId2',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
						rawTarget: '../queryTables/queryTable2.xml',
						resolvedTarget: 'xl/queryTables/queryTable2.xml',
						featureFamily: 'preservedQueryTable',
					},
				],
			},
		})

		const orphanIssues = result.issues.filter(
			(i) => i.rule === 'table-query-integrity' && i.details?.kind === 'orphan-query-table-part',
		)
		expect(result.passed).toBe(false)
		expect(orphanIssues).toHaveLength(1)
		expect(orphanIssues[0]?.refs).toEqual(['xl/queryTables/queryTable2.xml'])
	})

	test('detects stale worksheet-owned queryTable inventory', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.preservedXml = { partPath: 'xl/worksheets/sheet1.xml' }
		wb.connectionParts.push(
			{
				kind: 'queryTable',
				partPath: 'xl/queryTables/queryTable1.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
				sheetName: 'Sheet1',
				relationshipCount: 1,
				name: 'MissingPartQuery',
				connectionId: 1,
			},
			{
				kind: 'queryTable',
				partPath: 'xl/queryTables/queryTable2.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
				sheetName: 'Sheet1',
				relationshipCount: 1,
				name: 'WrongSheetQuery',
				connectionId: 2,
			},
			{
				kind: 'queryTable',
				partPath: 'xl/queryTables/queryTable3.xml',
				contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
				relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/unknownQuery',
				sheetName: 'Sheet1',
				relationshipCount: 1,
				name: 'WrongTypeQuery',
				connectionId: 3,
			},
		)

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/queryTables/queryTable2.xml',
						featureFamily: 'preservedQueryTable',
						ownerScope: 'worksheet',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
					},
					{
						path: 'xl/queryTables/queryTable3.xml',
						featureFamily: 'preservedQueryTable',
						ownerScope: 'worksheet',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet2.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet2.xml.rels',
						id: 'rId2',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
						rawTarget: '../queryTables/queryTable2.xml',
						resolvedTarget: 'xl/queryTables/queryTable2.xml',
						featureFamily: 'preservedQueryTable',
					},
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rId3',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable',
						rawTarget: '../queryTables/queryTable3.xml',
						resolvedTarget: 'xl/queryTables/queryTable3.xml',
						featureFamily: 'preservedQueryTable',
					},
				],
			},
		})

		expect(result.passed).toBe(false)
		expect(result.issues).toContainEqual(
			expect.objectContaining({
				rule: 'table-query-integrity',
				severity: 'error',
				refs: ['Sheet1', 'xl/queryTables/queryTable1.xml'],
				details: expect.objectContaining({
					kind: 'missing-worksheet-query-table-part',
					connectionPart: expect.objectContaining({
						partPath: 'xl/queryTables/queryTable1.xml',
						sheetName: 'Sheet1',
						connectionId: 1,
					}),
				}),
			}),
		)
		expect(result.issues).toContainEqual(
			expect.objectContaining({
				rule: 'table-query-integrity',
				severity: 'error',
				refs: ['Sheet1', 'xl/queryTables/queryTable2.xml'],
				details: expect.objectContaining({
					kind: 'worksheet-query-table-relationship-binding-mismatch',
					incomingRelationships: [
						expect.objectContaining({
							sourcePartPath: 'xl/worksheets/sheet2.xml',
							relationshipPartPath: 'xl/worksheets/_rels/sheet2.xml.rels',
							id: 'rId2',
						}),
					],
				}),
			}),
		)
		expect(result.issues).toContainEqual(
			expect.objectContaining({
				rule: 'table-query-integrity',
				severity: 'warning',
				refs: ['Sheet1', 'xl/queryTables/queryTable3.xml'],
				details: expect.objectContaining({
					kind: 'worksheet-query-table-relationship-type-mismatch',
					relationshipType:
						'http://schemas.openxmlformats.org/officeDocument/2006/relationships/unknownQuery',
				}),
			}),
		)
	})

	test('detects table package relationship binding mismatches and orphan table sidecars', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			partPath: 'xl/tables/table1.xml',
			sourcePartPath: 'xl/worksheets/sheet1.xml',
			sourceRelationshipPart: 'xl/worksheets/_rels/sheet1.xml.rels',
			sourceRelationshipId: 'rIdTable',
			sourceRelationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
			sourceRelationshipRawTarget: '../tables/table1.xml',
			sourceRelationshipResolvedTarget: 'xl/tables/table1.xml',
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Amount' }],
			hasHeaders: true,
			hasTotals: false,
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/tables/table1.xml',
						featureFamily: 'preservedTable',
						ownerScope: 'worksheet',
					},
					{
						path: 'xl/tables/table_custom.xml',
						featureFamily: 'preservedTable',
						ownerScope: 'worksheet',
					},
					{
						path: 'xl/tables/table_custom.bin',
						featureFamily: 'preservedTable',
						ownerScope: 'worksheet',
					},
					{
						path: 'xl/tables/_rels/table_custom.xml.rels',
						featureFamily: 'preservedTable',
						ownerScope: 'relationship-part',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdOther',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
						rawTarget: '../tables/table1.xml',
						resolvedTarget: 'xl/tables/table1.xml',
						featureFamily: 'preservedTable',
					},
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdOrphan',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
						rawTarget: '../tables/table_custom.xml',
						resolvedTarget: 'xl/tables/table_custom.xml',
						featureFamily: 'preservedTable',
					},
				],
			},
		})
		const packageIssues = result.issues.filter((i) => i.rule === 'table-package-integrity')
		expect(result.passed).toBe(false)
		expect(packageIssues.map((issue) => issue.details?.kind)).toEqual([
			'table-relationship-binding-mismatch',
			'orphan-table-part',
		])
		expect(packageIssues[0]?.refs).toEqual(['Sheet1!A1:B3', 'xl/tables/table1.xml'])
		expect(packageIssues[0]?.details).toMatchObject({
			table: {
				tableName: 'Sales',
				partPath: 'xl/tables/table1.xml',
				sourcePartPath: 'xl/worksheets/sheet1.xml',
				sourceRelationshipId: 'rIdTable',
			},
			incomingRelationships: [
				{
					sourcePartPath: 'xl/worksheets/sheet1.xml',
					id: 'rIdOther',
					resolvedTarget: 'xl/tables/table1.xml',
				},
			],
		})
		expect(packageIssues[1]?.severity).toBe('warning')
		expect(packageIssues[1]?.refs).toEqual(['xl/tables/table_custom.xml'])
	})

	test('detects orphan pivot cache, records, and table package sidecars', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/pivotCache/pivotCacheDefinition1.xml',
						featureFamily: 'preservedPivot',
						ownerScope: 'pivot',
					},
					{
						path: 'xl/pivotCache/pivotCacheRecords1.xml',
						featureFamily: 'preservedPivot',
						ownerScope: 'pivot',
					},
					{
						path: 'xl/pivotTables/pivotTable1.xml',
						featureFamily: 'preservedPivot',
						ownerScope: 'pivot',
					},
				],
				relationships: [],
			},
		})

		const kinds = result.issues
			.filter((issue) => issue.rule === 'pivot-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(kinds).toContain('orphan-pivot-cache-part')
		expect(kinds).toContain('orphan-pivot-cache-records-part')
		expect(kinds).toContain('orphan-pivot-table-part')
	})

	test('detects pivot cache relationship binding and records mismatches', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 7,
			relId: 'rIdPivotCache',
			recordCount: 2,
			recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml',
			records: {
				partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				declaredCount: 3,
				parsedCount: 3,
				preview: [
					{
						index: 0,
						values: [{ index: 0, kind: 'string', value: 'A' }],
					},
				],
				valueKindCounts: [{ kind: 'string', count: 1 }],
			},
			fields: [
				{ index: 0, name: 'Region' },
				{ index: 1, name: 'Amount' },
			],
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/pivotCache/pivotCacheDefinition1.xml',
						featureFamily: 'preservedPivot',
						ownerScope: 'pivot',
					},
					{
						path: 'xl/pivotCache/pivotCacheRecords1.xml',
						featureFamily: 'preservedPivot',
						ownerScope: 'pivot',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdOther',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
						rawTarget: 'pivotCache/pivotCacheDefinition1.xml',
						resolvedTarget: 'xl/pivotCache/pivotCacheDefinition1.xml',
						featureFamily: 'preservedPivot',
					},
					{
						sourcePartPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
						relationshipPartPath: 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
						id: 'rIdRecords',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
						rawTarget: 'pivotCacheRecords2.xml',
						resolvedTarget: 'xl/pivotCache/pivotCacheRecords2.xml',
						featureFamily: 'preservedPivot',
					},
				],
			},
		})

		const kinds = result.issues
			.filter((issue) => issue.rule === 'pivot-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(kinds).toContain('pivot-cache-workbook-relationship-binding-mismatch')
		expect(kinds).toContain('pivot-cache-records-relationship-binding-mismatch')
		expect(kinds).toContain('pivot-cache-record-count-mismatch')
		expect(kinds).toContain('pivot-cache-record-width-mismatch')
	})

	test('detects pivot cache definition and records payload path mismatches', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.pivotCaches.push(
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				cacheId: 1,
				recordsPartPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				records: {
					partPath: 'xl/pivotCache/pivotCacheRecords2.xml',
					declaredCount: 1,
					parsedCount: 1,
					preview: [],
					valueKindCounts: [],
				},
				fields: [],
			},
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition2.xml',
				cacheId: 2,
				records: {
					partPath: 'xl/pivotCache/pivotCacheRecords3.xml',
					declaredCount: 1,
					parsedCount: 1,
					preview: [],
					valueKindCounts: [],
				},
				fields: [],
			},
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition3.xml',
				cacheId: 3,
				recordsPartPath: 'xl/pivotCache/pivotCacheRecords4.xml',
				fields: [],
			},
		)

		const result = check(wb)
		const kinds = result.issues
			.filter((issue) => issue.rule === 'pivot-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(kinds).toContain('pivot-cache-records-part-path-mismatch')
		expect(kinds).toContain('pivot-cache-records-relationship-missing')
		expect(kinds).toContain('pivot-cache-records-unparsed')
	})

	test('detects duplicate and missing pivot cache ids and worksheet ownership', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.pivotCaches.push(
			{ partPath: 'xl/pivotCache/pivotCacheDefinition1.xml', cacheId: 1, fields: [] },
			{ partPath: 'xl/pivotCache/pivotCacheDefinition2.xml', cacheId: 1, fields: [] },
			{ partPath: 'xl/pivotCache/pivotCacheDefinition3.xml', fields: [] },
		)
		wb.pivotTables.push(
			{
				partPath: 'xl/pivotTables/pivotTable1.xml',
				sheetName: 'MissingSheet',
				name: 'PivotTable1',
				cacheId: 99,
				fields: [],
				rowFields: [],
				columnFields: [],
				pageFields: [],
				dataFields: [],
			},
			{
				partPath: 'xl/pivotTables/pivotTable2.xml',
				sheetName: 'Sheet1',
				name: 'PivotTable2',
				cacheId: 1,
				fields: [],
				rowFields: [],
				columnFields: [],
				pageFields: [],
				dataFields: [],
			},
			{
				partPath: 'xl/pivotTables/pivotTable3.xml',
				sheetName: 'Sheet1',
				name: 'PivotTable3',
				fields: [],
				rowFields: [],
				columnFields: [],
				pageFields: [],
				dataFields: [],
			},
		)

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/pivotTables/pivotTable1.xml',
						featureFamily: 'preservedPivot',
						ownerScope: 'pivot',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdPivotTable',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable',
						rawTarget: 'pivotTables/pivotTable1.xml',
						resolvedTarget: 'xl/pivotTables/pivotTable1.xml',
						featureFamily: 'preservedPivot',
					},
				],
			},
		})

		const kinds = result.issues
			.filter((issue) => issue.rule === 'pivot-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(kinds).toContain('pivot-cache-id-missing')
		expect(kinds).toContain('duplicate-pivot-cache-id')
		expect(kinds).toContain('pivot-table-cache-id-ambiguous')
		expect(kinds).toContain('pivot-table-cache-id-absent')
		expect(kinds).toContain('pivot-table-sheet-missing')
		expect(kinds).toContain('pivot-table-cache-id-missing')
		expect(kinds).toContain('pivot-table-worksheet-relationship-binding-mismatch')
	})

	test('detects missing workbook pivot cache relationships', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 1,
			fields: [],
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/pivotCache/pivotCacheDefinition1.xml',
						featureFamily: 'preservedPivot',
						ownerScope: 'workbook',
					},
				],
				relationships: [],
			},
		})

		const issue = result.issues.find(
			(entry) =>
				entry.rule === 'pivot-integrity' &&
				entry.details?.kind === 'pivot-cache-workbook-relationship-missing',
		)
		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('error')
		expect(issue?.refs).toEqual(['xl/pivotCache/pivotCacheDefinition1.xml'])
		expect(issue?.suggestedFix).toContain('workbook pivotCache relationship')
	})

	test('detects pivot cache source sheet and table integrity issues', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.tables.push({
			id: createTableId(),
			name: 'SalesTable',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Region' }, { name: 'Amount' }],
			hasHeaders: true,
			hasTotals: false,
		})
		wb.pivotCaches.push(
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				cacheId: 1,
				sourceType: 'worksheet',
				sourceSheet: 'MissingSheet',
				sourceRef: 'not-a-range',
				sourceName: 'MissingTable',
				fields: [],
			},
			{
				partPath: 'xl/pivotCache/pivotCacheDefinition2.xml',
				cacheId: 2,
				sourceType: 'worksheet',
				sourceSheet: 'OtherSheet',
				sourceName: 'SalesTable',
				fields: [],
			},
		)

		const result = check(wb)
		const kinds = result.issues
			.filter((issue) => issue.rule === 'pivot-source-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(kinds).toContain('pivot-cache-source-sheet-missing')
		expect(kinds).toContain('pivot-cache-source-ref-invalid')
		expect(kinds).toContain('pivot-cache-source-table-missing')
		expect(kinds).toContain('pivot-cache-source-table-sheet-mismatch')
	})

	test('detects slicer cache relationship and pivot binding mismatches', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 7,
			fields: [],
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 7,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_Region',
			pivotCacheId: 8,
			pivotTableNames: ['PivotTable1', 'MissingPivot'],
		})
		wb.slicers.push({
			partPath: 'xl/slicers/slicer1.xml',
			name: 'Region',
			cacheName: 'MissingSlicerCache',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/slicerCaches/slicerCache1.xml',
						featureFamily: 'preservedSlicer',
						ownerScope: 'workbook',
					},
					{
						path: 'xl/slicers/slicer2.xml',
						featureFamily: 'preservedSlicer',
						ownerScope: 'worksheet',
					},
				],
				relationships: [],
			},
		})

		const kinds = result.issues
			.filter((issue) => issue.rule === 'slicer-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(kinds).toContain('slicer-cache-name-missing')
		expect(kinds).toContain('slicer-cache-pivot-table-missing')
		expect(kinds).toContain('slicer-cache-pivot-cache-id-mismatch')
		expect(kinds).toContain('slicer-cache-workbook-relationship-binding-mismatch')
		expect(kinds).toContain('orphan-slicer-part')
	})

	test('detects timeline cache relationship and state binding mismatches', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 7,
			fields: [],
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 7,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.timelineCaches.push({
			partPath: 'xl/timelineCaches/timelineCache1.xml',
			name: 'Timeline_Date',
			pivotCacheId: 8,
			pivotTableNames: ['PivotTable1', 'MissingPivot'],
			state: {
				pivotCacheId: 9,
				filterPivotName: 'MissingFilterPivot',
			},
		})
		wb.timelines.push({
			partPath: 'xl/timelines/timeline1.xml',
			name: 'Date',
			cacheName: 'MissingTimelineCache',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/timelineCaches/timelineCache1.xml',
						featureFamily: 'preservedTimeline',
						ownerScope: 'workbook',
					},
					{
						path: 'xl/timelines/timeline2.xml',
						featureFamily: 'preservedTimeline',
						ownerScope: 'worksheet',
					},
				],
				relationships: [],
			},
		})

		const kinds = result.issues
			.filter((issue) => issue.rule === 'timeline-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(kinds).toContain('timeline-cache-name-missing')
		expect(kinds).toContain('timeline-state-pivot-cache-id-mismatch')
		expect(kinds).toContain('timeline-state-filter-pivot-missing')
		expect(kinds).toContain('timeline-cache-pivot-table-missing')
		expect(kinds).toContain('timeline-cache-pivot-cache-id-mismatch')
		expect(kinds).toContain('timeline-cache-workbook-relationship-binding-mismatch')
		expect(kinds).toContain('orphan-timeline-part')
	})

	test('detects slicer and timeline cache-to-UI package relationship loss', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.slicerCaches.push({
			partPath: 'xl/slicerCaches/slicerCache1.xml',
			name: 'Slicer_Region',
			pivotTableNames: [],
		})
		wb.slicers.push({
			partPath: 'xl/slicers/slicer1.xml',
			name: 'Region',
			cacheName: 'Slicer_Region',
		})
		wb.timelineCaches.push({
			partPath: 'xl/timelineCaches/timelineCache1.xml',
			name: 'Timeline_Date',
			pivotTableNames: [],
		})
		wb.timelines.push({
			partPath: 'xl/timelines/timeline1.xml',
			name: 'Date',
			cacheName: 'Timeline_Date',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/slicerCaches/slicerCache1.xml',
						featureFamily: 'preservedSlicer',
						ownerScope: 'slicer',
					},
					{
						path: 'xl/slicers/slicer1.xml',
						featureFamily: 'preservedSlicer',
						ownerScope: 'slicer',
					},
					{
						path: 'xl/timelineCaches/timelineCache1.xml',
						featureFamily: 'preservedTimeline',
						ownerScope: 'timeline',
					},
					{
						path: 'xl/timelines/timeline1.xml',
						featureFamily: 'preservedTimeline',
						ownerScope: 'timeline',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdSlicerCache',
						type: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
						rawTarget: 'slicerCaches/slicerCache1.xml',
						resolvedTarget: 'xl/slicerCaches/slicerCache1.xml',
					},
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdTimelineCache',
						type: 'http://schemas.microsoft.com/office/2011/relationships/timelineCache',
						rawTarget: 'timelineCaches/timelineCache1.xml',
						resolvedTarget: 'xl/timelineCaches/timelineCache1.xml',
					},
				],
			},
		})

		const slicerKinds = result.issues
			.filter((issue) => issue.rule === 'slicer-integrity')
			.map((issue) => issue.details?.kind)
		const timelineKinds = result.issues
			.filter((issue) => issue.rule === 'timeline-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(slicerKinds).toEqual([
			'slicer-cache-ui-relationship-binding-mismatch',
			'slicer-worksheet-relationship-binding-mismatch',
		])
		expect(timelineKinds).toEqual([
			'timeline-cache-ui-relationship-binding-mismatch',
			'timeline-worksheet-relationship-binding-mismatch',
		])
	})

	test('detects nonnumeric slicer and timeline orphan package sidecars', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/slicerCaches/cache_region.xml',
						featureFamily: 'preservedSlicer',
						ownerScope: 'workbook',
					},
					{
						path: 'xl/slicers/ui_region.xml',
						featureFamily: 'preservedSlicer',
						ownerScope: 'worksheet',
					},
					{
						path: 'xl/timelineCaches/cache_date.xml',
						featureFamily: 'preservedTimeline',
						ownerScope: 'workbook',
					},
					{
						path: 'xl/timelines/ui_date.xml',
						featureFamily: 'preservedTimeline',
						ownerScope: 'worksheet',
					},
					{
						path: 'xl/slicerCaches/_rels/cache_region.xml.rels',
						featureFamily: 'packageRelationships',
						ownerScope: 'relationship-part',
					},
					{
						path: 'xl/slicerCaches/cache_region.bin',
						featureFamily: 'preservedSlicer',
						ownerScope: 'workbook',
					},
					{
						path: 'xl/timelineCaches/_rels/cache_date.xml.rels',
						featureFamily: 'packageRelationships',
						ownerScope: 'relationship-part',
					},
					{
						path: 'xl/timelines/ui_date.bin',
						featureFamily: 'preservedTimeline',
						ownerScope: 'worksheet',
					},
				],
				relationships: [],
			},
		})

		const slicerKinds = result.issues
			.filter((issue) => issue.rule === 'slicer-integrity')
			.map((issue) => issue.details?.kind)
		const timelineKinds = result.issues
			.filter((issue) => issue.rule === 'timeline-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(slicerKinds).toEqual(['orphan-slicer-cache-part', 'orphan-slicer-part'])
		expect(timelineKinds).toEqual(['orphan-timeline-cache-part', 'orphan-timeline-part'])
	})

	test('surfaces stale pivot output and unsupported headless refresh signals', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 1,
			invalid: true,
			refreshOnLoad: true,
			saveData: false,
			enableRefresh: false,
			records: {
				partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				parsedCount: 10,
				preview: [],
				materializedCount: 5,
				materializedComplete: false,
				valueKindCounts: [],
			},
			fields: [],
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 1,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})

		const result = check(wb)
		const kinds = result.issues
			.filter((issue) => issue.rule === 'pivot-refresh-integrity')
			.map((issue) => issue.details?.kind)
		expect(result.passed).toBe(false)
		expect(kinds).toContain('pivot-cache-refresh-required')
		expect(kinds).toContain('pivot-headless-refresh-unsupported')
		expect(kinds).toContain('pivot-cache-refresh-disabled')
		expect(kinds).toContain('pivot-saved-output-stale-after-source-edit')
		expect(kinds).toContain('pivot-output-may-be-stale')
	})

	test('detects externalLink package binding mismatches', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rIdExternal',
			sourcePartPath: 'xl/workbook.xml',
			sourceRelationshipPart: 'xl/_rels/workbook.xml.rels',
			externalBookRelId: 'rIdBook',
			linkBindingStatus: 'externalBookRelId',
			target: '../sources/source.xlsx',
			targetMode: 'External',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
						ownerScope: 'external-link',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdOther',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
						rawTarget: 'externalLinks/externalLink1.xml',
						resolvedTarget: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
					},
					{
						sourcePartPath: 'xl/externalLinks/externalLink1.xml',
						relationshipPartPath: 'xl/externalLinks/_rels/externalLink1.xml.rels',
						id: 'rIdPath',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
						rawTarget: '../sources/source.xlsx',
						targetMode: 'External',
						featureFamily: 'preservedExternalLink',
					},
				],
			},
		})

		const externalIssues = result.issues.filter((i) => i.rule === 'external-link-integrity')
		expect(result.passed).toBe(false)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				message: expect.stringContaining('does not bind'),
				refs: ['xl/_rels/workbook.xml.rels#rIdExternal'],
				details: expect.objectContaining({
					kind: 'external-link-source-relationship-binding-mismatch',
				}),
			}),
		)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				message: expect.stringContaining('not claimed by externalLink metadata'),
				refs: ['xl/externalLinks/_rels/externalLink1.xml.rels#rIdPath'],
				details: expect.objectContaining({
					kind: 'orphan-external-link-path-relationship',
				}),
			}),
		)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				message: expect.stringContaining('rIdBook'),
				refs: ['xl/externalLinks/externalLink1.xml#rIdBook'],
				suggestedFix: expect.stringContaining('externalLinkPath relationship'),
				details: expect.objectContaining({
					kind: 'external-book-relationship-missing',
				}),
			}),
		)
	})

	test('detects orphan externalLink package sidecars and fallback binding risks', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.externalReferences.push('xl/externalLinks/externalLink2.xml')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink2.xml',
			relId: 'rIdExternal2',
			sourcePartPath: 'xl/workbook.xml',
			sourceRelationshipPart: 'xl/_rels/workbook.xml.rels',
			externalBookRelId: 'rIdBad',
			linkRelId: 'rIdPath',
			linkRelationshipPart: 'xl/externalLinks/_rels/externalLink2.xml.rels',
			linkRelationshipType:
				'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary',
			linkRelationshipRawTarget: 'library.xlsx',
			linkBindingStatus: 'fallbackPathRelationship',
			target: 'library.xlsx',
			targetMode: 'External',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
						ownerScope: 'external-link',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
					},
					{
						path: 'xl/externalLinks/externalLink2.xml',
						featureFamily: 'preservedExternalLink',
						ownerScope: 'external-link',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdExternal2',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
						rawTarget: 'externalLinks/externalLink2.xml',
						resolvedTarget: 'xl/externalLinks/externalLink2.xml',
						featureFamily: 'preservedExternalLink',
					},
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdOrphanExternal',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
						rawTarget: 'externalLinks/externalLink1.xml',
						resolvedTarget: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
					},
					{
						sourcePartPath: 'xl/externalLinks/externalLink2.xml',
						relationshipPartPath: 'xl/externalLinks/_rels/externalLink2.xml.rels',
						id: 'rIdBad',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
						rawTarget: '../worksheets/sheet1.xml',
						resolvedTarget: 'xl/worksheets/sheet1.xml',
						targetMode: 'Internal',
						featureFamily: 'worksheet',
					},
					{
						sourcePartPath: 'xl/externalLinks/externalLink2.xml',
						relationshipPartPath: 'xl/externalLinks/_rels/externalLink2.xml.rels',
						id: 'rIdPath',
						type: 'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary',
						rawTarget: 'library.xlsx',
						targetMode: 'External',
						featureFamily: 'preservedExternalLink',
					},
				],
			},
		})

		const externalIssues = result.issues.filter((i) => i.rule === 'external-link-integrity')
		expect(result.passed).toBe(false)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				message: expect.stringContaining('fallbackPathRelationship'),
				refs: [
					'xl/externalLinks/externalLink2.xml',
					'xl/externalLinks/_rels/externalLink2.xml.rels#rIdPath',
				],
				suggestedFix: expect.stringContaining('externalBook r:id'),
				details: expect.objectContaining({
					kind: 'external-link-binding-risk',
					externalBookRelId: 'rIdBad',
				}),
			}),
		)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				message: expect.stringContaining('rIdBad'),
				refs: ['xl/externalLinks/_rels/externalLink2.xml.rels#rIdBad'],
				suggestedFix: expect.stringContaining('externalBook r:id'),
				details: expect.objectContaining({
					kind: 'external-book-relationship-type-mismatch',
					partPath: 'xl/externalLinks/externalLink2.xml',
					externalBookRelId: 'rIdBad',
					linkRelId: 'rIdPath',
					actualType:
						'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
				}),
			}),
		)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				message: expect.stringContaining('rIdBad'),
				refs: ['xl/externalLinks/_rels/externalLink2.xml.rels#rIdBad'],
				details: expect.objectContaining({
					kind: 'external-book-relationship-target-mismatch',
					partPath: 'xl/externalLinks/externalLink2.xml',
					externalBookRelId: 'rIdBad',
					linkRelId: 'rIdPath',
					expectedRawTarget: 'library.xlsx',
					actualRawTarget: '../worksheets/sheet1.xml',
				}),
			}),
		)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				message: expect.stringContaining('TargetMode'),
				refs: ['xl/externalLinks/_rels/externalLink2.xml.rels#rIdBad'],
				details: expect.objectContaining({
					kind: 'external-book-relationship-target-mode-mismatch',
					externalBookRelId: 'rIdBad',
				}),
			}),
		)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				message: expect.stringContaining('not claimed'),
				refs: ['xl/externalLinks/externalLink1.xml'],
				details: expect.objectContaining({
					kind: 'orphan-external-link-part',
				}),
			}),
		)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				message: expect.stringContaining('not claimed by workbook externalReferences metadata'),
				refs: ['xl/_rels/workbook.xml.rels#rIdOrphanExternal'],
				suggestedFix: expect.stringContaining('Reconnect the workbook externalReference metadata'),
				details: expect.objectContaining({
					kind: 'orphan-external-link-relationship',
					relationship: expect.objectContaining({
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdOrphanExternal',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
						rawTarget: 'externalLinks/externalLink1.xml',
						resolvedTarget: 'xl/externalLinks/externalLink1.xml',
					}),
				}),
			}),
		)
	})

	test('surfaces package graph relationship target issues as check diagnostics', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/workbook.xml',
						featureFamily: 'workbook',
						ownerScope: 'workbook',
					},
					{
						path: 'xl/unknown/vendorBlob.xml',
						featureFamily: 'preservedOther',
						ownerScope: 'unknown',
						preservationPolicy: 'unknown-review-required',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdMissingSheet',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
						rawTarget: 'worksheets/missing.xml',
						resolvedTarget: 'xl/worksheets/missing.xml',
						featureFamily: 'worksheet',
					},
				],
			},
		})
		const packageIssues = result.issues.filter((i) => i.rule === 'package-graph-integrity')
		expect(result.passed).toBe(false)
		expect(packageIssues).toHaveLength(1)
		expect(packageIssues[0]?.details?.code).toBe('package_relationship_target')
		expect(packageIssues[0]?.refs).toEqual(['xl/_rels/workbook.xml.rels#rIdMissingSheet'])
		expect(packageIssues[0]?.suggestedFix).toContain('restore the referenced package part')
	})

	test('surfaces duplicate package relationship ids as check diagnostics', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{ path: 'xl/workbook.xml', featureFamily: 'workbook', ownerScope: 'workbook' },
					{
						path: 'xl/worksheets/sheet1.xml',
						featureFamily: 'worksheet',
						ownerScope: 'sheet',
					},
					{
						path: 'xl/worksheets/sheet2.xml',
						featureFamily: 'worksheet',
						ownerScope: 'sheet',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdSheet',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
						rawTarget: 'worksheets/sheet1.xml',
						resolvedTarget: 'xl/worksheets/sheet1.xml',
						featureFamily: 'worksheet',
					},
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdSheet',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
						rawTarget: 'worksheets/sheet2.xml',
						resolvedTarget: 'xl/worksheets/sheet2.xml',
						featureFamily: 'worksheet',
					},
				],
			},
		})
		const duplicateIssue = result.issues.find(
			(issue) => issue.details?.code === 'package_relationship_duplicate_id',
		)
		expect(result.passed).toBe(false)
		expect(duplicateIssue).toMatchObject({
			rule: 'package-graph-integrity',
			severity: 'error',
			refs: ['xl/_rels/workbook.xml.rels#rIdSheet'],
			suggestedFix: expect.stringContaining('unique'),
			details: {
				code: 'package_relationship_duplicate_id',
				sourcePartPath: 'xl/workbook.xml',
				relationshipPartPath: 'xl/_rels/workbook.xml.rels',
				relationshipId: 'rIdSheet',
				featureFamily: 'worksheet',
				actual: [
					expect.objectContaining({ target: 'worksheets/sheet1.xml' }),
					expect.objectContaining({ target: 'worksheets/sheet2.xml' }),
				],
			},
		})
	})

	test('surfaces package graph relationship source issues as check diagnostics', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/media/image1.png',
						featureFamily: 'preservedMedia',
						ownerScope: 'drawing',
					},
					{
						path: 'xl/drawings/_rels/missingDrawing.xml.rels',
						featureFamily: 'packageRelationships',
						ownerScope: 'relationship-part',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/drawings/missingDrawing.xml',
						relationshipPartPath: 'xl/drawings/_rels/missingDrawing.xml.rels',
						id: 'rIdImage',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
						rawTarget: '../media/image1.png',
						resolvedTarget: 'xl/media/image1.png',
						featureFamily: 'preservedMedia',
					},
				],
			},
		})

		const packageIssues = result.issues.filter((i) => i.rule === 'package-graph-integrity')
		expect(result.passed).toBe(false)
		expect(packageIssues).toHaveLength(1)
		expect(packageIssues[0]?.details).toMatchObject({
			code: 'package_relationship_source',
			sourcePartPath: 'xl/drawings/missingDrawing.xml',
			relationshipPartPath: 'xl/drawings/_rels/missingDrawing.xml.rels',
			relationshipId: 'rIdImage',
			featureFamily: 'preservedMedia',
		})
		expect(packageIssues[0]?.refs).toEqual(['xl/drawings/_rels/missingDrawing.xml.rels#rIdImage'])
		expect(packageIssues[0]?.suggestedFix).toContain('orphan relationship sidecar')
	})

	test('surfaces empty orphan relationship sidecars as check diagnostics', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/externalLinks/_rels/externalLink9.xml.rels',
						featureFamily: 'packageRelationships',
						ownerScope: 'relationship-part',
					},
				],
				relationships: [],
			},
		})

		const packageIssues = result.issues.filter((i) => i.rule === 'package-graph-integrity')
		expect(result.passed).toBe(false)
		expect(packageIssues).toHaveLength(1)
		expect(packageIssues[0]?.refs).toEqual(['xl/externalLinks/_rels/externalLink9.xml.rels'])
		expect(packageIssues[0]?.details).toMatchObject({
			code: 'package_relationship_source',
			sourcePartPath: 'xl/externalLinks/externalLink9.xml',
			relationshipPartPath: 'xl/externalLinks/_rels/externalLink9.xml.rels',
			featureFamily: 'packageRelationships',
			ownerScope: 'relationship-part',
		})
		expect(packageIssues[0]?.details?.relationshipId).toBeUndefined()
	})

	test('accepts empty relationship sidecars whose source part exists', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
						ownerScope: 'external-link',
					},
					{
						path: 'xl/externalLinks/_rels/externalLink1.xml.rels',
						featureFamily: 'packageRelationships',
						ownerScope: 'relationship-part',
					},
				],
				relationships: [],
			},
		})

		expect(
			result.issues.filter(
				(i) =>
					i.rule === 'package-graph-integrity' && i.details?.code === 'package_relationship_source',
			),
		).toHaveLength(0)
	})

	test('surfaces stale content type overrides as package graph diagnostics', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/workbook.xml',
						featureFamily: 'workbook',
						ownerScope: 'workbook',
						contentType:
							'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
					},
				],
				relationships: [],
				contentTypeOverrides: [
					{
						partPath: 'xl/comments1.xml',
						contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
					},
				],
			},
		})

		const packageIssues = result.issues.filter((i) => i.rule === 'package-graph-integrity')
		expect(result.passed).toBe(false)
		expect(packageIssues).toHaveLength(1)
		expect(packageIssues[0]?.refs).toEqual(['[Content_Types].xml', 'xl/comments1.xml'])
		expect(packageIssues[0]?.details).toMatchObject({
			code: 'package_content_type_override_target',
			partPath: 'xl/comments1.xml',
			contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
			expected: 'xl/comments1.xml',
			actual: undefined,
		})
		expect(packageIssues[0]?.suggestedFix).toContain('stale content type override')
	})

	test('surfaces content type override mismatches as package graph diagnostics', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/threadedComments/threadedComment1.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.ms-excel.threadedcomments+xml',
					},
				],
				relationships: [],
				contentTypeOverrides: [
					{
						partPath: 'xl/threadedComments/threadedComment1.xml',
						contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
					},
				],
			},
		})

		const packageIssues = result.issues.filter((i) => i.rule === 'package-graph-integrity')
		expect(result.passed).toBe(false)
		expect(packageIssues).toContainEqual(
			expect.objectContaining({
				refs: ['[Content_Types].xml', 'xl/threadedComments/threadedComment1.xml'],
				details: expect.objectContaining({
					code: 'package_content_type_override_mismatch',
					partPath: 'xl/threadedComments/threadedComment1.xml',
					expected: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
					actual: 'application/vnd.ms-excel.threadedcomments+xml',
					featureFamily: 'preservedThreadedComments',
				}),
			}),
		)
	})

	test('detects external link relationship mismatches and orphan package sidecars', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rIdExternal',
			sourcePartPath: 'xl/workbook.xml',
			sourceRelationshipPart: 'xl/_rels/workbook.xml.rels',
			sourceRelationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
			sourceRelationshipRawTarget: 'externalLinks/externalLink1.xml',
			sourceRelationshipResolvedTarget: 'xl/externalLinks/externalLink1.xml',
			externalBookRelId: 'rIdMissing',
			linkRelId: 'rIdPath',
			linkRelationshipPart: 'xl/externalLinks/_rels/externalLink1.xml.rels',
			linkRelationshipKind: 'externalLinkPath',
			linkBindingStatus: 'fallbackPathRelationship',
			linkRelationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
			linkRelationshipRawTarget: '../sources/source.xlsx',
			target: '../sources/source.xlsx',
			targetMode: 'External',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/workbook.xml',
						featureFamily: 'workbook',
						ownerScope: 'workbook',
					},
					{
						path: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
						ownerScope: 'workbook',
					},
					{
						path: 'xl/externalLinks/externalLink2.xml',
						featureFamily: 'preservedExternalLink',
						ownerScope: 'workbook',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdOther',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
						rawTarget: 'externalLinks/externalLink1.xml',
						resolvedTarget: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
					},
					{
						sourcePartPath: 'xl/externalLinks/externalLink1.xml',
						relationshipPartPath: 'xl/externalLinks/_rels/externalLink1.xml.rels',
						id: 'rIdPath',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
						rawTarget: '../sources/source.xlsx',
						targetMode: 'External',
						featureFamily: 'preservedExternalLink',
					},
				],
			},
		})
		const externalIssues = result.issues.filter((i) => i.rule === 'external-link-integrity')

		expect(result.passed).toBe(false)
		expect(
			externalIssues.some((issue) => issue.details?.kind === 'external-link-binding-risk'),
		).toBe(true)
		const sourceMismatch = externalIssues.find(
			(issue) => issue.details?.kind === 'external-link-source-relationship-binding-mismatch',
		)
		expect(sourceMismatch?.severity).toBe('error')
		expect(sourceMismatch?.refs).toEqual(['xl/_rels/workbook.xml.rels#rIdExternal'])
		expect(sourceMismatch?.details).toMatchObject({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relationshipId: 'rIdExternal',
			incomingRelationships: [{ id: 'rIdOther' }],
		})
		expect(
			externalIssues.some((issue) => issue.details?.kind === 'external-link-binding-risk'),
		).toBe(true)
		const orphanPart = externalIssues.find(
			(issue) => issue.details?.kind === 'orphan-external-link-part',
		)
		expect(orphanPart?.severity).toBe('warning')
		expect(orphanPart?.refs).toEqual(['xl/externalLinks/externalLink2.xml'])
	})

	test('detects external link package path target drift from metadata', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink1.xml',
			relId: 'rIdExternal',
			sourcePartPath: 'xl/workbook.xml',
			sourceRelationshipPart: 'xl/_rels/workbook.xml.rels',
			sourceRelationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
			sourceRelationshipRawTarget: 'externalLinks/externalLink1.xml',
			sourceRelationshipResolvedTarget: 'xl/externalLinks/externalLink1.xml',
			externalBookRelId: 'rIdPath',
			linkRelId: 'rIdPath',
			linkRelationshipPart: 'xl/externalLinks/_rels/externalLink1.xml.rels',
			linkRelationshipKind: 'externalLinkPath',
			linkBindingStatus: 'externalBookRelId',
			linkRelationshipType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
			linkRelationshipRawTarget: '../sources/source.xlsx',
			target: '../sources/source.xlsx',
			targetMode: 'External',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/workbook.xml',
						featureFamily: 'workbook',
						ownerScope: 'workbook',
					},
					{
						path: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
						ownerScope: 'workbook',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdExternal',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
						rawTarget: 'externalLinks/externalLink1.xml',
						resolvedTarget: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
					},
					{
						sourcePartPath: 'xl/externalLinks/externalLink1.xml',
						relationshipPartPath: 'xl/externalLinks/_rels/externalLink1.xml.rels',
						id: 'rIdPath',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
						rawType:
							'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
						rawTarget: '../sources/drifted.xlsx',
						targetMode: 'Internal',
						featureFamily: 'preservedExternalLink',
					},
				],
			},
		})
		const externalIssues = result.issues.filter((i) => i.rule === 'external-link-integrity')
		const typeMismatch = externalIssues.find(
			(issue) => issue.details?.kind === 'external-link-package-path-relationship-type-mismatch',
		)
		const targetMismatch = externalIssues.find(
			(issue) => issue.details?.kind === 'external-link-package-path-target-mismatch',
		)
		const targetModeMismatch = externalIssues.find(
			(issue) => issue.details?.kind === 'external-link-package-path-target-mode-mismatch',
		)

		expect(result.passed).toBe(false)
		expect(typeMismatch?.severity).toBe('warning')
		expect(typeMismatch?.refs).toEqual(['xl/externalLinks/_rels/externalLink1.xml.rels#rIdPath'])
		expect(typeMismatch?.details).toMatchObject({
			partPath: 'xl/externalLinks/externalLink1.xml',
			linkRelId: 'rIdPath',
			expectedType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
			actualType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
			actualRawType:
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
		})
		expect(targetMismatch?.severity).toBe('error')
		expect(targetMismatch?.refs).toEqual(['xl/externalLinks/_rels/externalLink1.xml.rels#rIdPath'])
		expect(targetMismatch?.details).toMatchObject({
			partPath: 'xl/externalLinks/externalLink1.xml',
			linkRelId: 'rIdPath',
			expectedRawTarget: '../sources/source.xlsx',
			actualRawTarget: '../sources/drifted.xlsx',
		})
		expect(targetModeMismatch?.severity).toBe('warning')
		expect(targetModeMismatch?.details).toMatchObject({
			expectedTargetMode: 'External',
			actualTargetMode: 'Internal',
		})
	})

	test('detects orphaned externalLink workbook metadata', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		wb.externalReferences.push('xl/externalLinks/externalLink1.xml')
		wb.externalReferenceDetails.push({
			partPath: 'xl/externalLinks/externalLink2.xml',
			relId: 'rIdExternal2',
			linkBindingStatus: 'externalBookRelId',
		})

		const result = check(wb)
		const externalIssues = result.issues.filter((i) => i.rule === 'external-link-integrity')

		expect(result.passed).toBe(false)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				refs: ['xl/externalLinks/externalLink1.xml'],
				details: expect.objectContaining({
					kind: 'external-reference-missing-detail',
				}),
			}),
		)
		expect(externalIssues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				refs: ['xl/externalLinks/externalLink2.xml'],
				details: expect.objectContaining({
					kind: 'orphan-external-link-metadata',
				}),
			}),
		)
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
			kind: 'conditional-format-priority-collision',
			priority: 1,
			left: { source: 'conditionalFormat', sqref: 'A1:A5', ruleType: 'expression' },
			right: { source: 'conditionalFormat', sqref: 'A3:A7', ruleType: 'cellIs' },
		})
	})

	test('detects non-positive conditional format priorities', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.conditionalFormats.push({
			sqref: 'E1',
			rules: [{ type: 'expression', priority: 0, formulas: ['E1>0'] }],
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'F1',
			priority: -1,
			type: 'dataBar',
			formulas: [],
		})

		const result = check(wb)
		const issues = result.issues.filter(
			(i) =>
				i.rule === 'conditional-format-integrity' &&
				i.details?.kind === 'conditional-format-nonpositive-priority',
		)

		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(2)
		expect(issues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				refs: ['Sheet1!E1'],
				details: expect.objectContaining({
					source: 'conditionalFormat',
					priority: 0,
					formatIndex: 0,
					ruleIndex: 0,
				}),
			}),
		)
		expect(issues).toContainEqual(
			expect.objectContaining({
				severity: 'warning',
				refs: ['Sheet1!F1'],
				details: expect.objectContaining({
					source: 'x14ConditionalFormat',
					priority: -1,
					formatIndex: 0,
					ruleType: 'dataBar',
				}),
			}),
		)
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
			kind: 'conditional-format-priority-collision',
			priority: 2,
			left: { source: 'conditionalFormat', sqref: 'B2:B5' },
			right: { source: 'x14ConditionalFormat', sqref: 'B4:B8', ruleType: 'dataBar' },
		})
	})

	test('detects broken legacy data validation references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.dataValidations.push(
			{
				sqref: 'MissingRange!C1',
				type: 'list',
				formula1: 'MissingList!A1:A5',
				formula2: '#REF!',
			},
			{
				sqref: 'A1',
				type: 'list',
				formula1: 'MissingTable[Name]',
			},
		)

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'data-validation-integrity')

		expect(result.passed).toBe(false)
		expect(issues.map((issue) => issue.details?.kind).sort()).toEqual([
			'data-validation-formula-deleted-reference',
			'data-validation-formula-missing-sheet',
			'data-validation-formula-missing-table',
			'data-validation-sqref-missing-sheet',
		])
		expect(issues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				refs: ['MissingRange!C1'],
				details: expect.objectContaining({
					kind: 'data-validation-sqref-missing-sheet',
					source: 'dataValidation',
					index: 0,
					field: 'sqref',
					reference: 'MissingRange!C1',
					missingSheet: 'MissingRange',
					token: 'MissingRange!C1',
					validationType: 'list',
				}),
			}),
		)
		expect(issues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				refs: ['Sheet1!A1'],
				details: expect.objectContaining({
					kind: 'data-validation-formula-missing-table',
					source: 'dataValidation',
					index: 1,
					field: 'formula1',
					reference: 'MissingTable[Name]',
					tableName: 'MissingTable',
					column: 'Name',
				}),
			}),
		)
	})

	test('detects broken legacy conditional format references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.conditionalFormats.push({
			sqref: 'MissingCfRange!A1:A5',
			rules: [
				{
					type: 'expression',
					priority: 1,
					formulas: ['MissingFormula!A1>0', '#REF!', 'MissingRules[Amount]>0'],
				},
				{
					type: 'dataBar',
					priority: 2,
					formulas: [],
					dataBar: { cfvo: [{ type: 'formula', value: 'SUM(#REF!)' }] },
				},
			],
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'conditional-format-integrity')

		expect(result.passed).toBe(false)
		expect(issues.map((issue) => issue.details?.kind).sort()).toEqual([
			'conditional-format-formula-deleted-reference',
			'conditional-format-formula-deleted-reference',
			'conditional-format-formula-missing-sheet',
			'conditional-format-formula-missing-table',
			'conditional-format-sqref-missing-sheet',
		])
		expect(issues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				refs: ['MissingCfRange!A1:A5'],
				details: expect.objectContaining({
					kind: 'conditional-format-sqref-missing-sheet',
					source: 'conditionalFormat',
					index: 0,
					field: 'sqref',
					reference: 'MissingCfRange!A1:A5',
					missingSheet: 'MissingCfRange',
				}),
			}),
		)
		expect(issues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				refs: ['MissingCfRange!A1:A5'],
				details: expect.objectContaining({
					kind: 'conditional-format-formula-deleted-reference',
					source: 'conditionalFormat',
					index: 0,
					ruleIndex: 1,
					ruleType: 'dataBar',
					field: 'rules[1].dataBar.cfvo[0].value',
					reference: 'SUM(#REF!)',
					error: '#REF!',
				}),
			}),
		)
	})

	test('detects broken x14 data validation sheet references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.x14DataValidations.push({
			index: 0,
			sqref: 'MissingRange!C1',
			type: 'list',
			formula1: 'MissingList!A1:A5',
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'data-validation-integrity')

		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(2)
		expect(issues.map((issue) => issue.details?.kind).sort()).toEqual([
			'x14-formula-missing-sheet',
			'x14-sqref-missing-sheet',
		])
		expect(issues.every((issue) => issue.severity === 'error')).toBe(true)
		expect(issues.find((issue) => issue.details?.field === 'formula1')?.details).toMatchObject({
			source: 'x14DataValidation',
			index: 0,
			missingSheet: 'MissingList',
			reference: 'MissingList!A1:A5',
		})
		expect(issues.find((issue) => issue.details?.field === 'sqref')?.refs).toEqual([
			'MissingRange!C1',
		])
	})

	test('detects broken x14 conditional format sheet references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'MissingCfRange!A1:A5',
			type: 'dataBar',
			formulas: ['MissingFormula!A1>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'MissingBar!A1' }] },
			iconSet: { cfvo: [{ type: 'formula', value: 'MissingIcon!A1' }] },
		})

		const result = check(wb)
		const issues = result.issues.filter(
			(i) =>
				i.rule === 'conditional-format-integrity' &&
				(i.details?.kind === 'x14-formula-missing-sheet' ||
					i.details?.kind === 'x14-sqref-missing-sheet'),
		)

		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(4)
		expect(issues.every((issue) => issue.severity === 'error')).toBe(true)
		expect(issues.map((issue) => issue.details?.field).sort()).toEqual([
			'dataBar.cfvo[0].value',
			'formulas[0]',
			'iconSet.cfvo[0].value',
			'sqref',
		])
		expect(issues[0]?.details).toMatchObject({
			source: 'x14ConditionalFormat',
			index: 0,
		})
	})

	test('detects malformed x14 sqref ranges', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.x14DataValidations.push({
			index: 0,
			sqref: 'A1:Bogus',
			type: 'list',
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: '#REF!',
			type: 'dataBar',
			formulas: [],
		})

		const result = check(wb)
		const invalidIssues = result.issues.filter(
			(i) =>
				(i.rule === 'data-validation-integrity' || i.rule === 'conditional-format-integrity') &&
				i.details?.kind === 'x14-sqref-invalid',
		)
		const missingSqrefIssues = result.issues.filter(
			(i) =>
				(i.rule === 'data-validation-integrity' || i.rule === 'conditional-format-integrity') &&
				i.details?.kind === 'x14-sqref-missing-sheet',
		)

		expect(result.passed).toBe(false)
		expect(invalidIssues).toHaveLength(2)
		expect(invalidIssues).toContainEqual(
			expect.objectContaining({
				rule: 'data-validation-integrity',
				severity: 'error',
				refs: ['Sheet1!A1:Bogus'],
				details: expect.objectContaining({
					source: 'x14DataValidation',
					field: 'sqref',
					reference: 'A1:Bogus',
					token: 'A1:Bogus',
				}),
			}),
		)
		expect(invalidIssues).toContainEqual(
			expect.objectContaining({
				rule: 'conditional-format-integrity',
				severity: 'error',
				details: expect.objectContaining({
					source: 'x14ConditionalFormat',
					field: 'sqref',
					reference: '#REF!',
					token: '#REF!',
				}),
			}),
		)
		expect(missingSqrefIssues).toHaveLength(0)
	})

	test('detects deleted references in x14 metadata formulas', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.x14DataValidations.push({
			index: 0,
			sqref: 'A1:A5',
			type: 'list',
			formula1: '#REF!',
			formula2: 'SUM(#REF!)',
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B1:B5',
			type: 'dataBar',
			formulas: ['#REF!'],
			dataBar: { cfvo: [{ type: 'formula', value: 'SUM(#REF!)' }] },
		})

		const result = check(wb)
		const deletedIssues = result.issues.filter(
			(i) =>
				(i.rule === 'data-validation-integrity' || i.rule === 'conditional-format-integrity') &&
				i.details?.kind === 'x14-formula-deleted-reference',
		)

		expect(result.passed).toBe(false)
		expect(deletedIssues).toHaveLength(4)
		expect(deletedIssues.every((issue) => issue.severity === 'error')).toBe(true)
		expect(deletedIssues.map((issue) => issue.details?.field).sort()).toEqual([
			'dataBar.cfvo[0].value',
			'formula1',
			'formula2',
			'formulas[0]',
		])
		expect(deletedIssues).toContainEqual(
			expect.objectContaining({
				rule: 'data-validation-integrity',
				refs: ['Sheet1!A1:A5'],
				details: expect.objectContaining({
					source: 'x14DataValidation',
					field: 'formula2',
					reference: 'SUM(#REF!)',
					error: '#REF!',
				}),
			}),
		)
		expect(deletedIssues).toContainEqual(
			expect.objectContaining({
				rule: 'conditional-format-integrity',
				refs: ['Sheet1!B1:B5'],
				details: expect.objectContaining({
					source: 'x14ConditionalFormat',
					field: 'dataBar.cfvo[0].value',
					reference: 'SUM(#REF!)',
					error: '#REF!',
				}),
			}),
		)
	})

	test('detects missing structured tables in x14 metadata formulas', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.x14DataValidations.push({
			index: 0,
			sqref: 'A1:A5',
			type: 'list',
			formula1: 'MissingList[Name]',
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'B1:B5',
			type: 'dataBar',
			formulas: ['SUM(MissingRules[Amount])>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'MissingBars[Value]' }] },
		})

		const result = check(wb)
		const validationIssues = result.issues.filter((i) => i.rule === 'data-validation-integrity')
		const formatIssues = result.issues.filter((i) => i.rule === 'conditional-format-integrity')

		expect(result.passed).toBe(false)
		expect(validationIssues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				refs: ['Sheet1!A1:A5'],
				details: expect.objectContaining({
					kind: 'x14-formula-missing-table',
					source: 'x14DataValidation',
					field: 'formula1',
					reference: 'MissingList[Name]',
					tableName: 'MissingList',
					column: 'Name',
				}),
			}),
		)
		expect(formatIssues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				refs: ['Sheet1!B1:B5'],
				details: expect.objectContaining({
					kind: 'x14-formula-missing-table',
					source: 'x14ConditionalFormat',
					field: 'formulas[0]',
					reference: 'SUM(MissingRules[Amount])>0',
					tableName: 'MissingRules',
					column: 'Amount',
				}),
			}),
		)
		expect(formatIssues).toContainEqual(
			expect.objectContaining({
				severity: 'error',
				refs: ['Sheet1!B1:B5'],
				details: expect.objectContaining({
					kind: 'x14-formula-missing-table',
					source: 'x14ConditionalFormat',
					field: 'dataBar.cfvo[0].value',
					reference: 'MissingBars[Value]',
					tableName: 'MissingBars',
					column: 'Value',
				}),
			}),
		)
	})

	test('detects duplicate x14 document-order indexes', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.x14DataValidations.push(
			{ index: 0, sqref: 'A1:A5', type: 'list' },
			{ index: 0, sqref: 'B1:B5', type: 'whole' },
		)
		s.x14ConditionalFormats.push(
			{ index: 3, sqref: 'C1:C5', type: 'dataBar', formulas: [] },
			{ index: 3, sqref: 'D1:D5', type: 'iconSet', formulas: [] },
		)

		const result = check(wb)

		expect(result.passed).toBe(false)
		expect(result.issues).toContainEqual(
			expect.objectContaining({
				rule: 'data-validation-integrity',
				severity: 'error',
				refs: ['Sheet1!A1:A5', 'Sheet1!B1:B5'],
				details: expect.objectContaining({
					kind: 'duplicate-x14-data-validation-index',
					source: 'x14DataValidation',
					index: 0,
					count: 2,
					sqrefs: ['A1:A5', 'B1:B5'],
				}),
			}),
		)
		expect(result.issues).toContainEqual(
			expect.objectContaining({
				rule: 'conditional-format-integrity',
				severity: 'error',
				refs: ['Sheet1!C1:C5', 'Sheet1!D1:D5'],
				details: expect.objectContaining({
					kind: 'duplicate-x14-conditional-format-index',
					source: 'x14ConditionalFormat',
					index: 3,
					count: 2,
					sqrefs: ['C1:C5', 'D1:D5'],
				}),
			}),
		)
	})

	test('detects deleted x14 entries that still carry live references', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.x14DataValidations.push({
			index: 0,
			sqref: 'A1',
			formula1: 'Sheet1!B1',
			formula2: '#REF!',
			deleted: true,
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'C1',
			formulas: ['Sheet1!D1>0'],
			dataBar: { cfvo: [{ type: 'formula', value: '#REF!' }] },
			deleted: true,
		})

		const result = check(wb)
		const dataValidationIssue = result.issues.find(
			(i) =>
				i.rule === 'data-validation-integrity' &&
				i.details?.kind === 'deleted-x14-data-validation-live-refs',
		)
		const conditionalFormatIssue = result.issues.find(
			(i) =>
				i.rule === 'conditional-format-integrity' &&
				i.details?.kind === 'deleted-x14-conditional-format-live-refs',
		)

		expect(result.passed).toBe(false)
		expect(dataValidationIssue?.severity).toBe('warning')
		expect(dataValidationIssue?.refs).toEqual(['Sheet1!A1'])
		expect(dataValidationIssue?.details).toMatchObject({
			source: 'x14DataValidation',
			index: 0,
			liveFields: ['sqref', 'formula1', 'formula2'],
		})
		expect(conditionalFormatIssue?.severity).toBe('warning')
		expect(conditionalFormatIssue?.refs).toEqual(['Sheet1!C1'])
		expect(conditionalFormatIssue?.details).toMatchObject({
			source: 'x14ConditionalFormat',
			index: 0,
			liveFields: ['sqref', 'formulas[0]', 'dataBar.cfvo[0].value'],
		})
	})

	test('detects ambiguous legacy and x14 conditional format overlap with missing priorities', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.conditionalFormats.push({
			sqref: 'D1:D5',
			rules: [{ type: 'expression', formulas: ['D1>0'] }],
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'D3:D7',
			type: 'dataBar',
			formulas: [],
		})

		const result = check(wb)
		const issue = result.issues.find(
			(i) =>
				i.rule === 'conditional-format-integrity' &&
				i.details?.kind === 'ambiguous-x14-legacy-overlap',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('warning')
		expect(issue?.refs).toEqual(['Sheet1!D1:D5', 'Sheet1!D3:D7'])
		expect(issue?.details).toMatchObject({
			left: { source: 'conditionalFormat', sqref: 'D1:D5', ruleType: 'expression' },
			right: { source: 'x14ConditionalFormat', sqref: 'D3:D7', ruleType: 'dataBar' },
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
			kind: 'duplicate-threaded-comment-id',
			partPath: 'xl/threadedComments/threadedComment1.xml',
			id: '{thread-1}',
			firstCommentIndex: 0,
			duplicateCommentIndex: 1,
		})
	})

	test('detects duplicate threaded comment ids across threaded comment parts', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		s1.threadedComments.push({
			ref: 'A1',
			text: 'Root',
			id: '{thread-1}',
			author: 'Alex',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})
		s2.threadedComments.push({
			ref: 'B2',
			text: 'Unexpected duplicate',
			id: '{thread-1}',
			author: 'Blair',
			partPath: 'xl/threadedComments/threadedComment2.xml',
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'threaded-comment-integrity')

		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(1)
		expect(issues[0]?.details).toEqual({
			kind: 'duplicate-threaded-comment-id-across-parts',
			id: '{thread-1}',
			firstPartPath: 'xl/threadedComments/threadedComment1.xml',
			duplicatePartPath: 'xl/threadedComments/threadedComment2.xml',
			firstSheetName: 'Sheet1',
			duplicateSheetName: 'Sheet2',
			firstCommentIndex: 0,
			duplicateCommentIndex: 0,
		})
		expect(issues[0]?.refs).toEqual(['Sheet1!A1', 'Sheet2!B2'])
		expect(issues[0]?.suggestedFix).toContain('workbook-unique')
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
			kind: 'missing-threaded-comment-parent-id',
			partPath: 'xl/threadedComments/threadedComment1.xml',
			commentIndex: 0,
			id: '{reply-1}',
			parentId: '{missing-root}',
		})
	})

	test('detects threaded comments that no longer reference valid worksheet cells', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push({
			ref: 'A0',
			text: 'Stale delete',
			id: '{thread-1}',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'threaded-comment-integrity')

		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('invalid cell reference "A0"')
		expect(issues[0]?.refs).toEqual(['Sheet1!A0'])
		expect(issues[0]?.details).toEqual({
			kind: 'invalid-threaded-comment-ref',
			partPath: 'xl/threadedComments/threadedComment1.xml',
			commentIndex: 0,
			ref: 'A0',
			id: '{thread-1}',
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
			kind: 'threaded-comment-unknown-person-id',
			partPath: 'xl/threadedComments/threadedComment1.xml',
			commentIndex: 0,
			id: '{thread-1}',
			personId: '{missing-person}',
		})
	})

	test('detects conflicting threaded comment authors for the same person id', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push(
			{
				ref: 'D4',
				text: 'Root',
				id: '{thread-1}',
				personId: '{person-1}',
				author: 'Ada',
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
			{
				ref: 'D5',
				text: 'Reply',
				id: '{thread-2}',
				parentId: '{thread-1}',
				personId: '{person-1}',
				author: 'Grace',
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
		)

		const result = check(wb)
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'threaded-comment-person-author-conflict',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('warning')
		expect(issue?.refs).toEqual(['Sheet1!D4', 'Sheet1!D5'])
		expect(issue?.details).toMatchObject({
			personId: '{person-1}',
			firstAuthor: 'Ada',
			duplicateAuthor: 'Grace',
			firstPartPath: 'xl/threadedComments/threadedComment1.xml',
			duplicatePartPath: 'xl/threadedComments/threadedComment1.xml',
			firstCommentIndex: 0,
			duplicateCommentIndex: 1,
			firstId: '{thread-1}',
			duplicateId: '{thread-2}',
		})
		expect(issue?.suggestedFix).toContain('one personId')
	})

	test('detects threaded comment missing ids and duplicate root refs', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push(
			{
				ref: 'A1',
				text: 'Root',
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
			{
				ref: 'A1',
				text: 'Second root',
				id: '{thread-2}',
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
		)

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'threaded-comment-integrity')

		expect(result.passed).toBe(false)
		expect(issues.some((issue) => issue.details?.kind === 'missing-threaded-comment-id')).toBe(true)
		expect(
			issues.some((issue) => issue.details?.kind === 'duplicate-threaded-comment-root-ref'),
		).toBe(true)
	})

	test('detects self-parented threaded comments', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push({
			ref: 'B2',
			text: 'Self parent',
			id: '{thread-1}',
			parentId: '{thread-1}',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const result = check(wb)
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'self-parented-threaded-comment',
		)

		expect(result.passed).toBe(false)
		expect(issue?.refs).toEqual(['Sheet1!B2'])
		expect(issue?.details).toMatchObject({
			partPath: 'xl/threadedComments/threadedComment1.xml',
			id: '{thread-1}',
			parentId: '{thread-1}',
		})
	})

	test('detects threaded comment parent id cycles', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push(
			{
				ref: 'B2',
				text: 'First reply',
				id: '{thread-1}',
				parentId: '{thread-2}',
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
			{
				ref: 'B3',
				text: 'Second reply',
				id: '{thread-2}',
				parentId: '{thread-1}',
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
		)

		const result = check(wb)
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'threaded-comment-parent-cycle',
		)

		expect(result.passed).toBe(false)
		expect(issue?.refs).toEqual(['Sheet1!B2', 'Sheet1!B3'])
		expect(issue?.details).toMatchObject({
			partPath: 'xl/threadedComments/threadedComment1.xml',
			cycleIds: ['{thread-1}', '{thread-2}'],
			commentIndexes: [0, 1],
		})
		expect(issue?.suggestedFix).toContain('parentId chains')
	})

	test('detects threaded comment package ownership ambiguities', () => {
		const wb = createWorkbook()
		const s1 = wb.addSheet('Sheet1')
		const s2 = wb.addSheet('Sheet2')
		s1.threadedComments.push({
			ref: 'A1',
			text: 'Root one',
			id: '{thread-1}',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})
		s2.threadedComments.push({
			ref: 'A1',
			text: 'Root two',
			id: '{thread-2}',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/threadedComments/threadedComment1.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.ms-excel.threadedcomments+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdThreaded1',
						type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
						rawTarget: '../threadedComments/threadedComment1.xml',
						resolvedTarget: 'xl/threadedComments/threadedComment1.xml',
					},
					{
						sourcePartPath: 'xl/worksheets/sheet2.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet2.xml.rels',
						id: 'rIdThreaded2',
						type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
						rawTarget: '../threadedComments/threadedComment1.xml',
						resolvedTarget: 'xl/threadedComments/threadedComment1.xml',
					},
				],
			},
		})
		const issues = result.issues.filter((i) => i.rule === 'threaded-comment-integrity')

		expect(result.passed).toBe(false)
		expect(
			issues.some((issue) => issue.details?.kind === 'threaded-comment-part-multiple-sheet-owners'),
		).toBe(true)
		expect(
			issues.some(
				(issue) => issue.details?.kind === 'threaded-comment-part-ambiguous-package-owner',
			),
		).toBe(true)
	})

	test('detects threaded comments that reference missing package parts', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push({
			ref: 'A1',
			text: 'Root',
			id: '{thread-1}',
			partPath: 'xl/threadedComments/missing.xml',
		})

		const result = check(wb, { packageGraph: { parts: [], relationships: [] } })
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'missing-threaded-comment-part',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('error')
		expect(issue?.refs).toEqual(['xl/threadedComments/missing.xml'])
	})

	test('detects threaded comment package owner drift from the model sheet', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.preservedXml = { partPath: 'xl/worksheets/sheet1.xml' }
		s.threadedComments.push({
			ref: 'A1',
			text: 'Root',
			id: '{thread-1}',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/threadedComments/threadedComment1.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.ms-excel.threadedcomments+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet2.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet2.xml.rels',
						id: 'rIdThreaded',
						type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
						rawTarget: '../threadedComments/threadedComment1.xml',
						resolvedTarget: 'xl/threadedComments/threadedComment1.xml',
					},
				],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'threaded-comment-sheet-owner-mismatch',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('error')
		expect(issue?.refs).toEqual([
			'Sheet1',
			'xl/worksheets/sheet1.xml',
			'xl/worksheets/sheet2.xml',
			'xl/threadedComments/threadedComment1.xml',
		])
		expect(issue?.details).toMatchObject({
			partPath: 'xl/threadedComments/threadedComment1.xml',
			sheetName: 'Sheet1',
			expectedWorksheetPartPath: 'xl/worksheets/sheet1.xml',
			actualWorksheetPartPath: 'xl/worksheets/sheet2.xml',
		})
		expect(issue?.suggestedFix).toContain('threadedComments relationship')
	})

	test('detects threaded comment relationships that resolve to missing package parts', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdThreaded',
						type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
						rawTarget: '../threadedComments/missing.xml',
						resolvedTarget: 'xl/threadedComments/missing.xml',
					},
				],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'threaded-comment-relationship-missing-target',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('error')
		expect(issue?.refs).toEqual(['xl/worksheets/_rels/sheet1.xml.rels#rIdThreaded'])
		expect(issue?.details?.relationship).toMatchObject({
			id: 'rIdThreaded',
			sourcePartPath: 'xl/worksheets/sheet1.xml',
			resolvedTarget: 'xl/threadedComments/missing.xml',
		})
	})

	test('detects orphaned threaded comment persons package parts', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/persons/person.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'workbook',
						contentType: 'application/vnd.ms-excel.person+xml',
					},
				],
				relationships: [],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'orphan-threaded-comment-persons-part',
		)

		expect(result.passed).toBe(false)
		expect(issue?.refs).toEqual(['xl/persons/person.xml'])
		expect(issue?.details).toMatchObject({
			partPath: 'xl/persons/person.xml',
			contentType: 'application/vnd.ms-excel.person+xml',
		})
	})

	test('detects threaded comment persons package parts when no comments use person ids', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push({
			ref: 'A1',
			text: 'Root',
			id: '{thread-1}',
			author: 'Ada',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/threadedComments/threadedComment1.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.ms-excel.threadedcomments+xml',
					},
					{
						path: 'xl/persons/person.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'workbook',
						contentType: 'application/vnd.ms-excel.person+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdThreaded',
						type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
						rawTarget: '../threadedComments/threadedComment1.xml',
						resolvedTarget: 'xl/threadedComments/threadedComment1.xml',
					},
				],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'orphan-threaded-comment-persons-part',
		)

		expect(result.passed).toBe(false)
		expect(issue?.refs).toEqual(['xl/persons/person.xml'])
		expect(issue?.details).toMatchObject({
			partPath: 'xl/persons/person.xml',
			threadedCommentsWithPersonIds: 0,
		})
	})

	test('detects ambiguous threaded comment persons package parts when comments use person ids', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push({
			ref: 'A1',
			text: 'Root',
			id: '{thread-1}',
			personId: '{person-1}',
			author: 'Ada',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/threadedComments/threadedComment1.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.ms-excel.threadedcomments+xml',
					},
					{
						path: 'xl/persons/person.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'workbook',
						contentType: 'application/vnd.ms-excel.person+xml',
					},
					{
						path: 'xl/persons/person2.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'workbook',
						contentType: 'application/vnd.ms-excel.person+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdThreaded',
						type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
						rawTarget: '../threadedComments/threadedComment1.xml',
						resolvedTarget: 'xl/threadedComments/threadedComment1.xml',
					},
				],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'ambiguous-threaded-comment-persons-parts',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('warning')
		expect(issue?.refs).toEqual(['xl/persons/person.xml', 'xl/persons/person2.xml'])
		expect(issue?.details).toMatchObject({
			threadedCommentPartPaths: ['xl/threadedComments/threadedComment1.xml'],
			personParts: [{ partPath: 'xl/persons/person.xml' }, { partPath: 'xl/persons/person2.xml' }],
		})
	})

	test('detects duplicate threaded comment person ids from package graph metadata', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.threadedComments.push({
			ref: 'A1',
			text: 'Root',
			id: '{thread-1}',
			personId: '{person-1}',
			author: 'Ada Duplicate',
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/threadedComments/threadedComment1.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.ms-excel.threadedcomments+xml',
					},
					{
						path: 'xl/persons/person.xml',
						featureFamily: 'preservedThreadedComments',
						ownerScope: 'workbook',
						contentType: 'application/vnd.ms-excel.person+xml',
						threadedCommentPersons: [
							{ id: '{person-1}', displayName: 'Ada', index: 0 },
							{ id: '{person-1}', displayName: 'Ada Duplicate', index: 1 },
						],
					},
				],
				relationships: [],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'threaded-comment-integrity' &&
				i.details?.kind === 'duplicate-threaded-comment-person-id',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('warning')
		expect(issue?.refs).toEqual(['xl/persons/person.xml'])
		expect(issue?.details).toMatchObject({
			partPath: 'xl/persons/person.xml',
			personId: '{person-1}',
			firstPersonIndex: 0,
			duplicatePersonIndex: 1,
			firstDisplayName: 'Ada',
			duplicateDisplayName: 'Ada Duplicate',
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
			kind: 'legacy-comment-vml-target-drift',
			ref: 'B2',
			expectedRow: 1,
			expectedColumn: 1,
			actualRow: 4,
			actualColumn: 3,
			shapeId: '_x0000_s1025',
		})
	})

	test('detects legacy comments that no longer reference valid worksheet cells', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('A0', {
			text: 'Stale delete',
			author: 'Ada',
			legacyDrawing: {
				shapeId: '_x0000_s1025',
				row: 0,
				column: 0,
				anchor: [0, 0, 0, 0, 1, 0, 2, 0],
			},
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'legacy-comment-drawing-integrity')

		expect(result.passed).toBe(false)
		expect(issues).toHaveLength(1)
		expect(issues[0]?.message).toContain('invalid cell reference "A0"')
		expect(issues[0]?.refs).toEqual(['Sheet1!A0'])
		expect(issues[0]?.details).toEqual({
			kind: 'legacy-comment-invalid-ref',
			ref: 'A0',
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
			kind: 'duplicate-legacy-comment-vml-shape-id',
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
			kind: 'legacy-comment-vml-anchor-invalid',
			ref: 'A1',
			anchor: [0, 0, 0, 0, 2, 0, -1, 0],
			shapeId: '_x0000_s1025',
		})
	})

	test('detects legacy comment VML shape id format, reversed anchors, and visibility conflicts', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('A1', {
			text: 'One',
			legacyDrawing: {
				shapeId: 'CommentShape1',
				row: 0,
				column: 0,
				style: 'position:absolute;visibility:hidden',
				visible: true,
				anchor: [2, 0, 4, 0, 1, 0, 3, 0],
			},
		})

		const result = check(wb)
		const issues = result.issues.filter((i) => i.rule === 'legacy-comment-drawing-integrity')

		expect(result.passed).toBe(false)
		expect(
			issues.some((issue) => issue.details?.kind === 'legacy-comment-vml-shape-id-format'),
		).toBe(true)
		expect(
			issues.some((issue) => issue.details?.kind === 'legacy-comment-vml-anchor-reversed'),
		).toBe(true)
		expect(
			issues.some((issue) => issue.details?.kind === 'legacy-comment-vml-visibility-conflict'),
		).toBe(true)
	})

	test('detects missing legacy comment VML sidecar relationships', () => {
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

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/comments1.xml',
						featureFamily: 'preservedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdComments',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
						rawTarget: '../comments1.xml',
						resolvedTarget: 'xl/comments1.xml',
					},
				],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'legacy-comment-drawing-integrity' &&
				i.details?.kind === 'missing-legacy-comment-vml-relationship',
		)

		expect(result.passed).toBe(false)
		expect(issue?.message).toContain('no VML drawing relationship')
		expect(issue?.details).toMatchObject({
			commentCount: 1,
			legacyDrawingCount: 1,
		})
	})

	test('detects orphaned classic comments package parts', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/comments1.xml',
						featureFamily: 'preservedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdComments',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
						rawTarget: '../comments1.xml',
						resolvedTarget: 'xl/comments1.xml',
					},
				],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'legacy-comment-drawing-integrity' &&
				i.details?.kind === 'orphan-legacy-comments-part',
		)

		expect(result.passed).toBe(false)
		expect(issue?.refs).toEqual(['xl/comments1.xml'])
	})

	test('detects extra orphaned classic comments parts when modeled comments exist', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.preservedXml = { partPath: 'xl/worksheets/sheet1.xml' }
		s.comments.set('A1', { text: 'Keep', author: 'Ada' })

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/comments1.xml',
						featureFamily: 'preservedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
					},
					{
						path: 'xl/comments2.xml',
						featureFamily: 'preservedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdComments',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
						rawTarget: '../comments1.xml',
						resolvedTarget: 'xl/comments1.xml',
					},
					{
						sourcePartPath: 'xl/worksheets/sheet2.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet2.xml.rels',
						id: 'rIdComments',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
						rawTarget: '../comments2.xml',
						resolvedTarget: 'xl/comments2.xml',
					},
				],
			},
		})
		const orphanIssues = result.issues.filter(
			(i) =>
				i.rule === 'legacy-comment-drawing-integrity' &&
				i.details?.kind === 'orphan-legacy-comments-part',
		)

		expect(result.passed).toBe(false)
		expect(orphanIssues).toHaveLength(1)
		expect(orphanIssues[0]?.refs).toEqual(['xl/comments2.xml'])
	})

	test('detects classic comments without VML that are missing comments relationships', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.comments.set('A1', { text: 'Plain note', author: 'Ada' })

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/comments1.xml',
						featureFamily: 'preservedComments',
						ownerScope: 'worksheet',
						contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
					},
				],
				relationships: [],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'legacy-comment-drawing-integrity' &&
				i.details?.kind === 'missing-legacy-comments-relationship',
		)

		expect(result.passed).toBe(false)
		expect(issue?.message).toContain('no comments relationship')
		expect(issue?.details).toMatchObject({
			commentCount: 1,
			legacyDrawingCount: 0,
		})
	})

	test('detects legacy comment VML note sidecars without comments XML', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.preservedXml = { partPath: 'xl/worksheets/sheet1.xml' }
		s.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
			source: 'vml',
			kind: 'shape',
			vmlObjectType: 'Note',
			vmlShapeId: '_x0000_s1025',
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{ path: 'xl/worksheets/sheet1.xml', featureFamily: 'worksheet' },
					{
						path: 'xl/drawings/vmlDrawing1.vml',
						featureFamily: 'preservedVml',
						ownerScope: 'worksheet',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdVml',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
						rawTarget: '../drawings/vmlDrawing1.vml',
						resolvedTarget: 'xl/drawings/vmlDrawing1.vml',
					},
				],
			},
		})
		const issue = result.issues.find(
			(i) =>
				i.rule === 'legacy-comment-drawing-integrity' &&
				i.details?.kind === 'legacy-comment-vml-without-comments-part',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('warning')
		expect(issue?.refs).toEqual(['xl/drawings/vmlDrawing1.vml'])
		expect(issue?.details).toMatchObject({
			noteShapeCount: 1,
			noteShapes: [
				{
					sheetName: 'Sheet1',
					objectIndex: 0,
					drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
					shapeId: '_x0000_s1025',
				},
			],
		})
	})

	test('detects duplicate raw legacy comment VML note shape ids from drawing inventory', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.preservedXml = { partPath: 'xl/worksheets/sheet1.xml' }
		s.drawingObjectRefs.push(
			{
				drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
				source: 'vml',
				kind: 'shape',
				vmlObjectType: 'Note',
				vmlShapeId: '_x0000_s1025',
			},
			{
				drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
				source: 'vml',
				kind: 'shape',
				vmlObjectType: 'Note',
				vmlShapeId: '_x0000_s1025',
			},
		)

		const result = check(wb, {
			packageGraph: {
				parts: [
					{ path: 'xl/worksheets/sheet1.xml', featureFamily: 'worksheet' },
					{
						path: 'xl/comments1.xml',
						featureFamily: 'preservedComments',
						ownerScope: 'worksheet',
					},
					{
						path: 'xl/drawings/vmlDrawing1.vml',
						featureFamily: 'preservedVml',
						ownerScope: 'worksheet',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdComments',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
						rawTarget: '../comments1.xml',
						resolvedTarget: 'xl/comments1.xml',
					},
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdVml',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
						rawTarget: '../drawings/vmlDrawing1.vml',
						resolvedTarget: 'xl/drawings/vmlDrawing1.vml',
					},
				],
			},
		})
		const duplicateIssue = result.issues.find(
			(i) =>
				i.rule === 'legacy-comment-drawing-integrity' &&
				i.details?.kind === 'duplicate-raw-legacy-comment-vml-shape-id',
		)

		expect(result.passed).toBe(false)
		expect(duplicateIssue?.refs).toEqual([
			'Sheet1#drawingObject0',
			'Sheet1#drawingObject1',
			'xl/drawings/vmlDrawing1.vml',
		])
		expect(duplicateIssue?.details).toMatchObject({
			drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
			shapeId: '_x0000_s1025',
			first: { sheetName: 'Sheet1', objectIndex: 0 },
			duplicate: { sheetName: 'Sheet1', objectIndex: 1 },
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

	test('detects external workbook chart series references without treating them as missing local sheets', () => {
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
		const issues = result.issues.filter((i) => i.rule === 'chart-series-integrity')
		expect(issues).toHaveLength(2)
		expect(issues.every((issue) => issue.severity === 'warning')).toBe(true)
		expect(issues.map((issue) => issue.details?.kind)).toEqual([
			'chart-series-external-reference',
			'chart-series-external-reference',
		])
		expect(issues[0]?.details).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			seriesIndex: 0,
			field: 'categoryRef',
			externalTarget: '[Book1.xlsx]',
		})
	})

	test('detects chart series references to missing structured tables', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Data')
		s.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Month' }, { name: 'Amount' }],
			hasHeaders: true,
			hasTotals: false,
		})
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Data',
			chartType: 'lineChart',
			series: [{ categoryRef: 'Sales[Month]', valueRef: 'MissingSales[Amount]' }],
		})

		const result = check(wb)
		const issue = result.issues.find(
			(i) =>
				i.rule === 'chart-series-integrity' &&
				i.details?.kind === 'chart-series-missing-table-reference',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('warning')
		expect(issue?.refs).toEqual(['xl/charts/chart1.xml#series0'])
		expect(issue?.details).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			seriesIndex: 0,
			field: 'valueRef',
			tableName: 'MissingSales',
			ownerSheet: 'Data',
		})
	})

	test('detects external refs in table, validation, and conditional metadata', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.tables.push({
			id: createTableId(),
			name: 'Sales',
			sheetId: s.id,
			partPath: 'xl/tables/table1.xml',
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Amount', formula: '[Budget.xlsx]Data!B2' }],
			hasHeaders: true,
			hasTotals: false,
		})
		s.dataValidations.push({
			sqref: 'B2:B5',
			type: 'list',
			formula1: '[Lists.xlsx]Valid!$A$1:$A$5',
		})
		s.x14DataValidations.push({
			index: 0,
			sqref: "'[Lists.xlsx]Valid'!$C$1",
			type: 'whole',
		})
		s.conditionalFormats.push({
			sqref: 'C2:C5',
			rules: [{ type: 'expression', priority: 1, formulas: ['[Rules.xlsx]Sheet1!A1>0'] }],
		})
		s.x14ConditionalFormats.push({
			index: 0,
			sqref: 'D2:D5',
			type: 'dataBar',
			formulas: [],
			dataBar: { cfvo: [{ type: 'formula', value: '[Rules.xlsx]Sheet1!B1' }] },
		})

		const result = check(wb)
		const tableIssue = result.issues.find(
			(issue) =>
				issue.rule === 'table-integrity' &&
				issue.details?.kind === 'tableColumn-external-reference',
		)
		const validationIssues = result.issues.filter(
			(issue) =>
				issue.rule === 'data-validation-integrity' &&
				(issue.details?.kind === 'dataValidation-external-reference' ||
					issue.details?.kind === 'x14DataValidation-external-reference'),
		)
		const conditionalIssues = result.issues.filter(
			(issue) =>
				issue.rule === 'conditional-format-integrity' &&
				(issue.details?.kind === 'conditionalFormat-external-reference' ||
					issue.details?.kind === 'x14ConditionalFormat-external-reference'),
		)

		expect(result.passed).toBe(false)
		expect(tableIssue?.severity).toBe('warning')
		expect(tableIssue?.refs).toEqual(['Sheet1!A1:B3'])
		expect(tableIssue?.details).toMatchObject({
			tableName: 'Sales',
			columnName: 'Amount',
			field: 'columns[1].formula',
			externalTarget: '[Budget.xlsx]',
		})
		expect(validationIssues).toHaveLength(2)
		expect(validationIssues.map((issue) => issue.details?.externalTarget).sort()).toEqual([
			'[Lists.xlsx]',
			'[Lists.xlsx]',
		])
		expect(conditionalIssues).toHaveLength(2)
		expect(conditionalIssues.map((issue) => issue.details?.externalTarget).sort()).toEqual([
			'[Rules.xlsx]',
			'[Rules.xlsx]',
		])
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

	test('detects chart parts claimed by both worksheet and chartsheet owners', () => {
		const wb = createWorkbook()
		wb.addSheet('Data')
		wb.chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Data',
			chartType: 'lineChart',
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
		const issue = result.issues.find(
			(i) =>
				i.rule === 'chart-part-ownership' &&
				i.details?.kind === 'chart-worksheet-chartsheet-owner-ambiguity',
		)

		expect(result.passed).toBe(false)
		expect(issue?.severity).toBe('warning')
		expect(issue?.details).toEqual({
			kind: 'chart-worksheet-chartsheet-owner-ambiguity',
			partPath: 'xl/charts/chart1.xml',
			ownerSheet: 'Data',
			ownerChartSheet: 'Chart 1',
			chartType: 'lineChart',
		})
	})

	test('detects drawing chart and image package integrity issues', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.imageRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			relId: 'rIdImg',
			targetPath: 'xl/media/image1.png',
		})
		s.drawingObjectRefs.push(
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				source: 'drawingml',
				kind: 'graphicFrame',
				name: 'Shared Shape',
				relIds: ['rIdChart'],
			},
			{
				drawingPartPath: 'xl/drawings/vmlDrawing1.vml',
				source: 'vml',
				kind: 'shape',
				name: 'Shared Shape',
			},
		)

		const result = check(wb, {
			packageGraph: {
				parts: [
					{ path: 'xl/worksheets/sheet1.xml', featureFamily: 'worksheet' },
					{ path: 'xl/workbook.xml', featureFamily: 'workbook' },
					{
						path: 'xl/drawings/drawing1.xml',
						featureFamily: 'preservedDrawing',
						ownerScope: 'drawing',
						contentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
					},
					{
						path: 'xl/drawings/vmlDrawing1.vml',
						featureFamily: 'preservedVml',
						ownerScope: 'drawing',
					},
					{
						path: 'xl/drawings/drawing2.xml',
						featureFamily: 'preservedDrawing',
						ownerScope: 'unknown',
					},
					{
						path: 'xl/drawings/drawing3.xml',
						featureFamily: 'preservedDrawing',
						ownerScope: 'workbook',
					},
					{ path: 'xl/charts/chart1.xml', featureFamily: 'preservedChart' },
					{
						path: 'xl/media/image2.png',
						featureFamily: 'preservedMedia',
						contentType: 'image/png',
					},
				],
				relationships: [
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdDrawing',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
						rawTarget: '../drawings/drawing1.xml',
						resolvedTarget: 'xl/drawings/drawing1.xml',
					},
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdVml',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
						rawTarget: '../drawings/vmlDrawing1.vml',
						resolvedTarget: 'xl/drawings/vmlDrawing1.vml',
					},
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdWrongDrawing',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
						rawTarget: '../drawings/vmlDrawing1.vml',
						resolvedTarget: 'xl/drawings/vmlDrawing1.vml',
					},
					{
						sourcePartPath: 'xl/worksheets/sheet1.xml',
						relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
						id: 'rIdWrongVml',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
						rawTarget: '../drawings/drawing1.xml',
						resolvedTarget: 'xl/drawings/drawing1.xml',
					},
					{
						sourcePartPath: 'xl/workbook.xml',
						relationshipPartPath: 'xl/_rels/workbook.xml.rels',
						id: 'rIdBadOwner',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
						rawTarget: 'drawings/drawing3.xml',
						resolvedTarget: 'xl/drawings/drawing3.xml',
					},
					{
						sourcePartPath: 'xl/drawings/drawing1.xml',
						relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
						id: 'rIdImg',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
						rawTarget: '../media/image2.png',
						resolvedTarget: 'xl/media/image2.png',
					},
					{
						sourcePartPath: 'xl/drawings/drawing1.xml',
						relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
						id: 'rIdMissingMedia',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
						rawTarget: '../media/missing.png',
						resolvedTarget: 'xl/media/missing.png',
					},
					{
						sourcePartPath: 'xl/drawings/drawing1.xml',
						relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
						id: 'rIdChart',
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
						rawTarget: '../media/image2.png',
						resolvedTarget: 'xl/media/image2.png',
					},
				],
			},
		})
		const drawingIssues = result.issues.filter((i) => i.rule === 'drawing-integrity')
		const kinds = new Set(drawingIssues.map((issue) => issue.details?.kind))

		expect(result.passed).toBe(false)
		expect(kinds.has('image-media-target-mismatch')).toBe(true)
		expect(kinds.has('image-media-relationship-missing-target')).toBe(true)
		expect(kinds.has('drawing-chart-target-type-mismatch')).toBe(true)
		expect(kinds.has('orphan-drawing-part')).toBe(true)
		expect(kinds.has('drawing-missing-worksheet-chartsheet-owner')).toBe(true)
		expect(kinds.has('vml-drawingml-ownership-ambiguity')).toBe(true)
		expect(kinds.has('drawing-relationship-target-type-mismatch')).toBe(true)
		expect(kinds.has('vml-drawing-relationship-target-type-mismatch')).toBe(true)
		expect(
			drawingIssues.find((issue) => issue.details?.kind === 'image-media-target-mismatch')?.details,
		).toMatchObject({
			expectedTargetPath: 'xl/media/image1.png',
			actualTargetPath: 'xl/media/image2.png',
			relationshipId: 'rIdImg',
		})
	})

	test('detects malformed and reversed drawing anchors', () => {
		const wb = createWorkbook()
		const s = wb.addSheet('Sheet1')
		s.imageRefs.push(
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				relId: 'rIdBadImage',
				targetPath: 'xl/media/image1.png',
				anchor: {
					kind: 'oneCell',
					from: { row: 1, col: 0, rowOff: -1 },
				},
			},
			{
				drawingPartPath: 'xl/drawings/drawing1.xml',
				relId: 'rIdReversedImage',
				targetPath: 'xl/media/image2.png',
				anchor: {
					kind: 'twoCell',
					from: { row: 4, col: 3 },
					to: { row: 4, col: 2 },
				},
			},
		)
		s.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			source: 'drawingml',
			kind: 'shape',
			anchor: {
				kind: 'absolute',
				x: 0,
				y: -1,
			},
		})

		const result = check(wb, {
			packageGraph: {
				parts: [
					{
						path: 'xl/drawings/drawing1.xml',
						featureFamily: 'preservedDrawing',
					},
				],
				relationships: [],
			},
		})
		const anchorIssues = result.issues.filter(
			(issue) =>
				issue.rule === 'drawing-integrity' &&
				typeof issue.details?.kind === 'string' &&
				issue.details.kind.includes('anchor'),
		)

		expect(result.passed).toBe(false)
		expect(anchorIssues.map((issue) => issue.details?.kind).sort()).toEqual([
			'drawing-object-anchor-invalid',
			'image-anchor-invalid',
			'image-anchor-reversed',
		])
		expect(
			anchorIssues.find((issue) => issue.details?.kind === 'image-anchor-reversed')?.severity,
		).toBe('warning')
	})

	test('detects chart style and color sidecar orphan and mismatch issues', () => {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')

		const result = check(wb, {
			packageGraph: {
				parts: [
					{ path: 'xl/charts/chart1.xml', featureFamily: 'preservedChart' },
					{ path: 'xl/drawings/drawing1.xml', featureFamily: 'preservedDrawing' },
					{ path: 'xl/charts/style1.xml', featureFamily: 'preservedChartStyle' },
					{ path: 'xl/charts/style2.xml', featureFamily: 'preservedChartStyle' },
					{ path: 'xl/charts/colors1.xml', featureFamily: 'preservedChartColor' },
					{ path: 'xl/charts/colors2.xml', featureFamily: 'preservedChartColor' },
				],
				relationships: [
					{
						sourcePartPath: 'xl/charts/chart1.xml',
						relationshipPartPath: 'xl/charts/_rels/chart1.xml.rels',
						id: 'rIdStyle',
						type: 'http://schemas.microsoft.com/office/2011/relationships/chartStyle',
						rawTarget: 'colors1.xml',
						resolvedTarget: 'xl/charts/colors1.xml',
					},
					{
						sourcePartPath: 'xl/charts/chart1.xml',
						relationshipPartPath: 'xl/charts/_rels/chart1.xml.rels',
						id: 'rIdColor',
						type: 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle',
						rawTarget: 'colors1.xml',
						resolvedTarget: 'xl/charts/colors1.xml',
					},
					{
						sourcePartPath: 'xl/drawings/drawing1.xml',
						relationshipPartPath: 'xl/drawings/_rels/drawing1.xml.rels',
						id: 'rIdStyleOwner',
						type: 'http://schemas.microsoft.com/office/2011/relationships/chartStyle',
						rawTarget: '../charts/style1.xml',
						resolvedTarget: 'xl/charts/style1.xml',
					},
					{
						sourcePartPath: 'xl/charts/chart1.xml',
						relationshipPartPath: 'xl/charts/_rels/chart1.xml.rels',
						id: 'rIdMissingColor',
						type: 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle',
						rawTarget: 'colors9.xml',
						resolvedTarget: 'xl/charts/colors9.xml',
					},
				],
			},
		})
		const issues = result.issues.filter((i) => i.rule === 'chart-package-integrity')
		const kinds = new Set(issues.map((issue) => issue.details?.kind))

		expect(result.passed).toBe(false)
		expect(kinds.has('chart-style-color-target-mismatch')).toBe(true)
		expect(kinds.has('chart-style-color-owner-mismatch')).toBe(true)
		expect(kinds.has('orphan-chart-style-part')).toBe(true)
		expect(kinds.has('orphan-chart-color-part')).toBe(true)
		expect(kinds.has('chart-style-color-relationship-missing-target')).toBe(true)
		expect(issues.find((issue) => issue.details?.kind === 'orphan-chart-style-part')?.refs).toEqual(
			['xl/charts/style2.xml'],
		)
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

	test('reports structured-reference precedents through defined names', () => {
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
		wb.definedNames.set('ScoreValues', 'Scores[Score]')
		s.cells.set(4, 0, { value: EMPTY, formula: 'SUM(ScoreValues)', styleId: SID })

		const result = trace(wb, 'Sheet1', 'A5')
		expect(result.ok).toBe(true)
		if (!result.ok) throw new Error(result.error.message)

		expect(result.value.precedents).toHaveLength(1)
		expect(result.value.precedents[0]?.ref).toBe('B2:B3')
	})
})
