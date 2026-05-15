import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { DEFAULT_STYLE_ID, indexToColumn, parseA1 } from '@ascend/core'
import { errorValue, numberValue, type Operation, stringValue } from '@ascend/schema'
import {
	analyzeMutationJournalExactness,
	buildMutationJournal,
	classifyMutationJournalIssue,
	classifyMutationJournalIssues,
	classifyMutationJournalOperationPrimarySurface,
	classifyMutationJournalOperationSurfaces,
	classifyMutationJournalSurface,
	MUTATION_JOURNAL_EXACTNESS_MATRIX,
	MUTATION_JOURNAL_OPERATION_SURFACE_RULES,
	MUTATION_JOURNAL_REASON_DESCRIPTIONS,
	type MutationJournal,
	type MutationJournalOperationName,
	type MutationJournalReasonCode,
	type MutationJournalSurface,
	unavailableMutationJournal,
} from './journal.ts'
import { AscendWorkbook } from './workbook.ts'

const REQUIRED_JOURNAL_SURFACES: readonly MutationJournalSurface[] = [
	'cells',
	'formulas',
	'formula-bindings',
	'shared-formulas',
	'dynamic-arrays',
	'legacy-arrays',
	'data-tables',
	'spills',
	'tables',
	'defined-names',
	'comments',
	'hyperlinks',
	'data-validations',
	'conditional-formats',
	'auto-filters',
	'merged-cells',
	'row-layout',
	'column-layout',
	'page-setup',
	'sheet-layout',
	'x14-metadata',
	'drawings',
	'charts',
	'pivot-caches',
	'workbook-metadata',
	'package-parts',
]

const REQUIRED_JOURNAL_OPERATIONS: readonly MutationJournalOperationName[] = [
	'setCells',
	'setFormula',
	'fillFormula',
	'clearRange',
	'insertRows',
	'deleteRows',
	'insertCols',
	'deleteCols',
	'addSheet',
	'deleteSheet',
	'renameSheet',
	'moveSheet',
	'createTable',
	'appendRows',
	'sortRange',
	'mergeCells',
	'unmergeCells',
	'setColWidth',
	'setRowHeight',
	'setComment',
	'setHyperlink',
	'setNumberFormat',
	'setDefinedName',
	'deleteDefinedName',
	'setStyle',
	'freezePane',
	'deleteComment',
	'deleteHyperlink',
	'setDataValidation',
	'deleteDataValidation',
	'setAutoFilter',
	'clearAutoFilter',
	'setSheetProtection',
	'setTabColor',
	'hideSheet',
	'hideRows',
	'hideCols',
	'copySheet',
	'setConditionalFormat',
	'deleteConditionalFormat',
	'setPageSetup',
	'setPrintArea',
	'copyRange',
	'moveRange',
	'groupRows',
	'groupCols',
	'setRichText',
	'setWorkbookProperties',
	'setDocumentProperties',
	'setWorkbookView',
	'setCalcSettings',
	'setTheme',
	'setWorkbookProtection',
	'deleteTable',
	'renameTable',
	'resizeTable',
	'setTableColumn',
	'setTableStyle',
	'replaceImage',
	'insertImage',
	'deleteImage',
	'setDrawingText',
	'setThreadedComment',
	'setChartSeriesSource',
	'setPivotCache',
	'setPivotFieldItem',
	'setSlicerCacheItem',
	'setTimelineRange',
	'setSparklineGroup',
	'setAdvancedFilter',
	'setConnectionRefresh',
	'rewriteExternalLink',
]

const SYNTHETIC_REPRESENTATIVE_OPS = new Set(['unsupported', 'preserved-package-part'])

describe('mutation journal exactness model', () => {
	test('taxonomy covers every workbook edit surface with stable reason codes', () => {
		const surfaces = new Set(MUTATION_JOURNAL_EXACTNESS_MATRIX.map((rule) => rule.surface))
		expect(surfaces.size).toBe(MUTATION_JOURNAL_EXACTNESS_MATRIX.length)
		expect([...surfaces].sort()).toEqual([...REQUIRED_JOURNAL_SURFACES].sort())

		const reasonCodes = new Set(
			Object.keys(MUTATION_JOURNAL_REASON_DESCRIPTIONS) as MutationJournalReasonCode[],
		)
		const matrixReasonCodes = new Set(
			MUTATION_JOURNAL_EXACTNESS_MATRIX.flatMap((rule) => rule.lossReasons),
		)
		for (const rule of MUTATION_JOURNAL_EXACTNESS_MATRIX) {
			expect(classifyMutationJournalSurface(rule.surface)).toBe(rule)
			expect(rule.constraints.length).toBeGreaterThan(0)
			expect(rule.representativeOps.length).toBeGreaterThan(0)
			if (rule.exactness === 'exact') expect(rule.lossReasons).toEqual([])
			else expect(rule.lossReasons.length).toBeGreaterThan(0)
			for (const reason of rule.lossReasons) expect(reasonCodes.has(reason)).toBe(true)
		}
		for (const reason of reasonCodes) expect(matrixReasonCodes.has(reason)).toBe(true)
	})

	test('operation surface taxonomy covers every public operation and matrix representative', () => {
		const operationNames = Object.keys(
			MUTATION_JOURNAL_OPERATION_SURFACE_RULES,
		) as MutationJournalOperationName[]
		expect(new Set(REQUIRED_JOURNAL_OPERATIONS).size).toBe(REQUIRED_JOURNAL_OPERATIONS.length)
		expect([...operationNames].sort()).toEqual([...REQUIRED_JOURNAL_OPERATIONS].sort())

		const matrixSurfaces = new Set(MUTATION_JOURNAL_EXACTNESS_MATRIX.map((rule) => rule.surface))
		for (const op of operationNames) {
			const rule = MUTATION_JOURNAL_OPERATION_SURFACE_RULES[op]
			expect(rule.surfaces).toContain(rule.primarySurface)
			expect(new Set(rule.surfaces).size).toBe(rule.surfaces.length)
			expect(classifyMutationJournalOperationPrimarySurface(op)).toBe(rule.primarySurface)
			expect(classifyMutationJournalOperationSurfaces(op)).toEqual(rule.surfaces)
			for (const surface of rule.surfaces) expect(matrixSurfaces.has(surface)).toBe(true)
		}

		for (const rule of MUTATION_JOURNAL_EXACTNESS_MATRIX) {
			for (const op of rule.representativeOps) {
				if (SYNTHETIC_REPRESENTATIVE_OPS.has(op)) continue
				const operationRule =
					MUTATION_JOURNAL_OPERATION_SURFACE_RULES[op as MutationJournalOperationName]
				expect(operationRule.surfaces).toContain(rule.surface)
			}
		}
	})

	test('classifies known journal failures into the shared surface and reason vocabulary', () => {
		const packageIssue = unavailableMutationJournal('Partial workbook cannot prove package parts', [
			'package:xl/worksheets/sheet1.xml',
		]).issues[0]
		if (!packageIssue) throw new Error('missing package issue')
		expect(classifyMutationJournalIssue(packageIssue)).toEqual({
			surface: 'package-parts',
			reason: 'journal-unavailable',
			exactness: 'lossy',
			publicInverse: 'none',
		})

		const formulaBindingIssue = {
			code: 'LOSSY_INVERSE' as const,
			message: 'Formula binding metadata for Sheet1!A1 cannot be restored with public operations',
			refs: ['Sheet1!A1'],
		}
		expect(classifyMutationJournalIssue(formulaBindingIssue)).toEqual({
			surface: 'formula-bindings',
			reason: 'formula-binding-metadata',
			exactness: 'lossy',
			publicInverse: 'none',
		})

		const x14Issue = {
			code: 'LOSSY_INVERSE' as const,
			message:
				'Transferred x14 data validation metadata on Sheet1!A1 cannot be restored with public operations',
			refs: ['Sheet1!x14Validation:A1:A1:0'],
		}
		expect(classifyMutationJournalIssue(x14Issue)).toEqual({
			surface: 'x14-metadata',
			reason: 'x14-metadata',
			exactness: 'lossy',
			publicInverse: 'none',
		})

		const partialIssue = unavailableMutationJournal(
			'Partial workbook cannot prove an exact rollback journal',
			undefined,
			{ reason: 'partial-workbook' },
		).issues[0]
		if (!partialIssue) throw new Error('missing partial workbook issue')
		expect(classifyMutationJournalIssue(partialIssue)).toEqual({
			surface: 'package-parts',
			reason: 'partial-workbook',
			exactness: 'lossy',
			publicInverse: 'none',
		})
	})

	test('generated lossy journals classify every issue into an allowed surface reason', () => {
		const journals = [
			lossyDataValidationDefaultJournal(),
			lossyDataValidationDeleteOrderJournal(),
			lossyDataValidationDuplicateDeleteJournal(),
			lossyDataValidationDuplicateMoveJournal(),
			lossyLegacyCommentDrawingJournal(),
			lossySetCommentLegacyDrawingJournal(),
			lossyCreatedLayoutJournal(),
			lossyThreadedCommentSelectorJournal(),
			lossyDrawingSelectorJournal(),
			lossyChartSeriesUnsetJournal(),
			lossyPivotCacheUnsetJournal(),
			lossyConditionalFormatOrderJournal(),
			lossyConditionalFormatReplacementOrderJournal(),
			lossyConditionalFormatDeleteOrderJournal(),
			lossyConditionalFormatDuplicateDeleteJournal(),
			lossyConditionalFormatDuplicateMoveJournal(),
			lossyMergeDuplicateMoveJournal(),
			lossyAutoFilterJournal(),
			lossyPageSetupJournal(),
			lossyX14TransferJournal(),
			lossyOverlappingMoveRangeJournal(),
			lossyStyleRegistryJournal(),
			lossyDeletedSheetJournal(),
		]

		for (const journal of journals) {
			expect(journal.supported).toBe(true)
			expect(journal.exact).toBe(false)
			expect(journal.issues.length).toBeGreaterThan(0)
			for (const entry of journal.entries) {
				for (const issue of entry.issues) {
					expect(issue.surface).toBeDefined()
					expect(issue.reason).toBeDefined()
					expect(classifyMutationJournalIssue(issue)).toMatchObject({
						surface: issue.surface,
						reason: issue.reason,
					})
				}
			}
			for (const issue of journal.issues) {
				expect(issue.surface).toBeDefined()
				expect(issue.reason).toBeDefined()
				expect(classifyMutationJournalIssue(issue)).toMatchObject({
					surface: issue.surface,
					reason: issue.reason,
				})
			}
			for (const classification of classifyMutationJournalIssues(journal.issues)) {
				const rule = classifyMutationJournalSurface(classification.surface)
				expect(rule.lossReasons).toContain(classification.reason)
			}
		}
	})

	test('classifies autofilter losses by column extension and sort reasons', () => {
		const reasons = classifyMutationJournalIssues(lossyAutoFilterJournal().issues).map(
			(classification) => classification.reason,
		)
		expect(reasons).toEqual([
			'auto-filter-column-metadata',
			'auto-filter-extension-metadata',
			'auto-filter-sort-metadata',
		])
	})

	test('classifies data validation delete order as metadata order', () => {
		const journal = lossyDataValidationDeleteOrderJournal()
		expect(journal.exact).toBe(false)
		expect(journal.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message: 'Data-validation order on Sheet1 cannot be restored exactly with public operations',
			surface: 'data-validations',
			reason: 'metadata-order',
			refs: ['Sheet1!A1:A1'],
		})
	})

	test('classifies duplicate data validation deletion as metadata duplicate', () => {
		const journal = lossyDataValidationDuplicateDeleteJournal()
		expect(journal.exact).toBe(false)
		expect(journal.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message:
				'Duplicate data validation metadata on Sheet1 cannot be restored exactly with public operations',
			surface: 'data-validations',
			reason: 'metadata-duplicate',
			refs: ['Sheet1!A1:A1'],
		})
	})

	test('classifies duplicate data validation moves as metadata duplicate', () => {
		const journal = lossyDataValidationDuplicateMoveJournal()
		expect(journal.exact).toBe(false)
		expect(journal.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message:
				'Moved duplicate data validations on Sheet1!A1 cannot be restored exactly with public operations',
			surface: 'data-validations',
			reason: 'metadata-duplicate',
			refs: ['Sheet1!A1:A1'],
		})
	})

	test('classifies conditional format replacement order as metadata order', () => {
		const journal = lossyConditionalFormatReplacementOrderJournal()
		expect(journal.exact).toBe(false)
		expect(journal.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message:
				'Conditional-format order on Sheet1 cannot be restored exactly with public operations',
			surface: 'conditional-formats',
			reason: 'metadata-order',
			refs: ['Sheet1!A1:A2'],
		})
	})

	test('classifies conditional format delete order as metadata order', () => {
		const journal = lossyConditionalFormatDeleteOrderJournal()
		expect(journal.exact).toBe(false)
		expect(journal.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message:
				'Conditional-format order on Sheet1 cannot be restored exactly with public operations',
			surface: 'conditional-formats',
			reason: 'metadata-order',
			refs: ['Sheet1!A1:A2'],
		})
	})

	test('classifies duplicate conditional format deletion as metadata duplicate', () => {
		const journal = lossyConditionalFormatDuplicateDeleteJournal()
		expect(journal.exact).toBe(false)
		expect(journal.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message:
				'Duplicate conditional format metadata on Sheet1 cannot be restored exactly with public operations',
			surface: 'conditional-formats',
			reason: 'metadata-duplicate',
			refs: ['Sheet1!A1:A1'],
		})
	})

	test('classifies duplicate conditional format moves as metadata duplicate', () => {
		const journal = lossyConditionalFormatDuplicateMoveJournal()
		expect(journal.exact).toBe(false)
		expect(journal.issues).toContainEqual({
			code: 'LOSSY_INVERSE',
			message:
				'Moved duplicate conditional formats on Sheet1!A1 cannot be restored exactly with public operations',
			surface: 'conditional-formats',
			reason: 'metadata-duplicate',
			refs: ['Sheet1!A1:A1'],
		})
	})

	test('classifies overlapping moveRange cell inverse loss without matrix violation', () => {
		const analysis = analyzeMutationJournalExactness(lossyOverlappingMoveRangeJournal())
		expect(analysis).toMatchObject({
			supported: true,
			exact: false,
			issueCount: 1,
			surfaces: ['cells'],
			reasons: ['operation-unsupported'],
			hasLossyInverse: true,
			hasUnsupportedOperation: false,
			hasMatrixViolation: false,
		})
		expect(analysis.issues[0]).toMatchObject({
			code: 'LOSSY_INVERSE',
			surface: 'cells',
			reason: 'operation-unsupported',
			allowedByMatrix: true,
		})
	})

	test('classifies missing-sheet journal preimage failures as sheet topology', () => {
		const ops: readonly Operation[] = [
			{ op: 'setTabColor', sheet: 'Missing', color: 'FF0000' },
			{ op: 'setSheetProtection', sheet: 'Missing', options: { sort: true } },
			{ op: 'moveSheet', sheet: 'Missing', position: 0 },
			{ op: 'setRowHeight', sheet: 'Missing', row: 1, height: 24 },
			{ op: 'setColWidth', sheet: 'Missing', col: 1, width: 16 },
			{ op: 'hideSheet', sheet: 'Missing', hidden: true },
			{ op: 'hideRows', sheet: 'Missing', at: 0, count: 1, hidden: true },
			{ op: 'hideCols', sheet: 'Missing', at: 0, count: 1, hidden: true },
			{ op: 'groupRows', sheet: 'Missing', from: 0, to: 1 },
			{ op: 'groupCols', sheet: 'Missing', from: 0, to: 1 },
		]

		for (const op of ops) {
			const wb = AscendWorkbook.create()
			const analysis = analyzeMutationJournalExactness(
				buildMutationJournal(wb.getWorkbookModel(), [op]),
			)
			expect(analysis, op.op).toMatchObject({
				supported: true,
				exact: false,
				issueCount: 1,
				surfaces: ['sheet-layout'],
				reasons: ['sheet-topology'],
				hasMatrixViolation: false,
			})
			expect(analysis.issues[0], op.op).toMatchObject({
				code: 'UNSUPPORTED_VALUE',
				surface: 'sheet-layout',
				reason: 'sheet-topology',
				allowedByMatrix: true,
			})
		}
	})

	test('classifies setComment legacy drawing loss without changing the v1 vocabulary', () => {
		const classified = classifyMutationJournalIssues(lossySetCommentLegacyDrawingJournal().issues)
		expect(classified).toContainEqual({
			surface: 'comments',
			reason: 'legacy-comment-drawing',
			exactness: 'conditional',
			publicInverse: 'conditional',
		})
	})

	test('analyzes journal exactness into agent-readable surfaces and reasons', () => {
		const exactWorkbook = AscendWorkbook.create()
		applyExact(exactWorkbook, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
		])
		const exact = analyzeMutationJournalExactness(
			applyJournal(exactWorkbook, [
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
			]),
		)
		expect(exact).toMatchObject({
			supported: true,
			exact: true,
			issueCount: 0,
			issues: [],
			operationSurfaces: [
				'cells',
				'data-tables',
				'dynamic-arrays',
				'formula-bindings',
				'formulas',
				'legacy-arrays',
				'shared-formulas',
				'spills',
			],
			primaryOperationSurfaces: ['cells'],
			surfaces: [],
			reasons: [],
			hasMatrixViolation: false,
		})

		const lossy = analyzeMutationJournalExactness(lossyAutoFilterJournal())
		expect(lossy).toMatchObject({
			supported: true,
			exact: false,
			issueCount: 3,
			operationSurfaces: ['auto-filters'],
			primaryOperationSurfaces: ['auto-filters'],
			surfaces: ['auto-filters'],
			reasons: [
				'auto-filter-column-metadata',
				'auto-filter-extension-metadata',
				'auto-filter-sort-metadata',
			],
			hasLossyInverse: true,
			hasUnsupportedOperation: false,
			hasUnavailableJournal: false,
			hasJournalBuildFailure: false,
			hasMatrixViolation: false,
		})
		expect(lossy.issues.map((issue) => issue.allowedByMatrix)).toEqual([true, true, true])

		const unavailable = analyzeMutationJournalExactness(
			unavailableMutationJournal('Partial workbook cannot prove package parts', [
				'package:xl/workbook.xml',
			]),
		)
		expect(unavailable).toMatchObject({
			supported: false,
			exact: false,
			issueCount: 1,
			operationSurfaces: [],
			primaryOperationSurfaces: [],
			surfaces: ['package-parts'],
			reasons: ['journal-unavailable'],
			hasUnavailableJournal: true,
			hasMatrixViolation: false,
		})
	})

	test('classifies unsupported operation journals by workbook surface', () => {
		const cases: readonly {
			readonly op: Operation
			readonly surface: MutationJournalSurface
		}[] = [
			{ op: { op: 'appendRows', table: 'Sales', rows: [['East']] }, surface: 'tables' },
			{
				op: {
					op: 'copyRange',
					sheet: 'Sheet1',
					source: 'A1',
					target: 'B1',
					mode: 'transpose',
				} as Operation,
				surface: 'cells',
			},
			{
				op: {
					op: 'replaceImage',
					sheet: 'Sheet1',
					contentBase64: 'AA==',
					contentType: 'image/png',
					imageIndex: 0,
				},
				surface: 'drawings',
			},
			{
				op: {
					op: 'setAdvancedFilter',
					sheet: 'Sheet1',
					filterIndex: 0,
					range: 'A1:B10',
				},
				surface: 'auto-filters',
			},
			{
				op: {
					op: 'setPivotFieldItem',
					fieldIndex: 0,
					itemIndex: 0,
					hidden: true,
				},
				surface: 'pivot-caches',
			},
			{
				op: {
					op: 'setSparklineGroup',
					sheet: 'Sheet1',
					groupIndex: 0,
					range: 'A1:A5',
				},
				surface: 'x14-metadata',
			},
			{
				op: {
					op: 'rewriteExternalLink',
					relId: 'rId1',
					newTarget: '../external.xlsx',
				},
				surface: 'package-parts',
			},
		]

		for (const entry of cases) {
			const wb = AscendWorkbook.create()
			const analysis = analyzeMutationJournalExactness(
				buildMutationJournal(wb.getWorkbookModel(), [entry.op]),
			)
			expect(analysis).toMatchObject({
				supported: false,
				exact: false,
				issueCount: 1,
				surfaces: [entry.surface],
				reasons: ['operation-unsupported'],
				hasUnsupportedOperation: true,
				hasMatrixViolation: false,
			})
		}
	})

	test('representative edit operations obey the exactness taxonomy', () => {
		const cases: readonly {
			readonly surface: MutationJournalSurface
			readonly run: () => MutationJournal
			readonly exact: boolean
			readonly issue?: {
				readonly surface: MutationJournalSurface
				readonly reason: MutationJournalReasonCode
			}
		}[] = [
			{
				surface: 'cells',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])
					return applyJournal(wb, [
						{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
					])
				},
			},
			{
				surface: 'formulas',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [
						{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
						{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
					])
					return applyJournal(wb, [
						{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*3' },
					])
				},
			},
			{
				surface: 'shared-formulas',
				exact: false,
				issue: { surface: 'shared-formulas', reason: 'formula-binding-metadata' },
				run: () => {
					const wb = AscendWorkbook.create()
					const sheet = wb.getWorkbookModel().getSheet('Sheet1')
					if (!sheet) throw new Error('missing sheet')
					sheet.cells.set(0, 0, {
						value: numberValue(4),
						formula: 'B1*2',
						styleId: DEFAULT_STYLE_ID,
						formulaInfo: {
							kind: 'shared',
							sharedIndex: 'exactness-model',
							isMaster: true,
							masterRef: 'A1',
							ref: 'A1:A2',
						},
					})
					return applyJournal(wb, [
						{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] },
					])
				},
			},
			{
				surface: 'dynamic-arrays',
				exact: false,
				issue: { surface: 'dynamic-arrays', reason: 'formula-binding-metadata' },
				run: () => {
					const wb = AscendWorkbook.create()
					const sheet = wb.getWorkbookModel().getSheet('Sheet1')
					if (!sheet) throw new Error('missing sheet')
					sheet.cells.set(0, 0, {
						value: numberValue(1),
						formula: 'SEQUENCE(2)',
						styleId: DEFAULT_STYLE_ID,
						formulaInfo: { kind: 'dynamicArray', metadataIndex: 1 },
					})
					return applyJournal(wb, [
						{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] },
					])
				},
			},
			{
				surface: 'legacy-arrays',
				exact: false,
				issue: { surface: 'legacy-arrays', reason: 'formula-binding-metadata' },
				run: () => {
					const wb = AscendWorkbook.create()
					const sheet = wb.getWorkbookModel().getSheet('Sheet1')
					if (!sheet) throw new Error('missing sheet')
					sheet.cells.set(0, 0, {
						value: numberValue(1),
						formula: 'A1:A2*2',
						styleId: DEFAULT_STYLE_ID,
						formulaInfo: { kind: 'array', ref: 'A1:A2' },
					})
					return buildMutationJournal(wb.getWorkbookModel(), [
						{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] },
					])
				},
			},
			{
				surface: 'data-tables',
				exact: false,
				issue: { surface: 'data-tables', reason: 'formula-binding-metadata' },
				run: () => {
					const wb = AscendWorkbook.create()
					const sheet = wb.getWorkbookModel().getSheet('Sheet1')
					if (!sheet) throw new Error('missing sheet')
					sheet.cells.set(2, 2, {
						value: numberValue(10),
						formula: null,
						styleId: DEFAULT_STYLE_ID,
						formulaInfo: { kind: 'dataTable', ref: 'C3:C5', dtr: true, r1: 'A1' },
					})
					return applyJournal(wb, [
						{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'C3', value: 7 }] },
					])
				},
			},
			{
				surface: 'spills',
				exact: false,
				issue: { surface: 'spills', reason: 'formula-binding-metadata' },
				run: () => {
					const wb = AscendWorkbook.create()
					const sheet = wb.getWorkbookModel().getSheet('Sheet1')
					if (!sheet) throw new Error('missing sheet')
					sheet.cells.set(1, 0, {
						value: numberValue(2),
						formula: null,
						styleId: DEFAULT_STYLE_ID,
						formulaInfo: {
							kind: 'spill',
							anchorRef: 'Sheet1!A1',
							ref: 'A1:A3',
							isAnchor: false,
						},
					})
					return applyJournal(wb, [
						{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 7 }] },
					])
				},
			},
			{
				surface: 'tables',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [
						{
							op: 'setCells',
							sheet: 'Sheet1',
							updates: [
								{ ref: 'A1', value: 'Name' },
								{ ref: 'A2', value: 'West' },
							],
						},
						{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:A2', name: 'Sales', hasHeaders: true },
					])
					return applyJournal(wb, [{ op: 'renameTable', table: 'Sales', newName: 'Revenue' }])
				},
			},
			{
				surface: 'defined-names',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$A$1' }])
					return applyJournal(wb, [{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$B$1' }])
				},
			},
			{
				surface: 'comments',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [
						{ op: 'setComment', sheet: 'Sheet1', ref: 'A1', text: 'before', author: 'Ada' },
					])
					return applyJournal(wb, [
						{ op: 'setComment', sheet: 'Sheet1', ref: 'A1', text: 'after', author: 'Ada' },
					])
				},
			},
			{
				surface: 'hyperlinks',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [
						{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', url: 'https://before.example' },
					])
					return applyJournal(wb, [
						{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', url: 'https://after.example' },
					])
				},
			},
			{
				surface: 'data-validations',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [
						{
							op: 'setDataValidation',
							sheet: 'Sheet1',
							range: 'A1:A2',
							rule: { type: 'whole', operator: 'between', formula1: '1', formula2: '9' },
						},
					])
					return applyJournal(wb, [
						{
							op: 'setDataValidation',
							sheet: 'Sheet1',
							range: 'A1:A2',
							rule: { type: 'whole', operator: 'greaterThan', formula1: '3' },
						},
					])
				},
			},
			{
				surface: 'conditional-formats',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [
						{
							op: 'setConditionalFormat',
							sheet: 'Sheet1',
							range: 'A1:A2',
							rule: { type: 'expression', formula: 'A1>0', priority: 1 },
						},
					])
					return applyJournal(wb, [
						{
							op: 'setConditionalFormat',
							sheet: 'Sheet1',
							range: 'A1:A2',
							rule: { type: 'expression', formula: 'A1>5', priority: 1 },
						},
					])
				},
			},
			{
				surface: 'merged-cells',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [{ op: 'mergeCells', sheet: 'Sheet1', range: 'A1:B1' }])
					return applyJournal(wb, [{ op: 'unmergeCells', sheet: 'Sheet1', range: 'A1:B1' }])
				},
			},
			{
				surface: 'auto-filters',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [
						{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:B10' },
						{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:B10', column: 0, values: ['Open'] },
					])
					return applyJournal(wb, [{ op: 'clearAutoFilter', sheet: 'Sheet1' }])
				},
			},
			{
				surface: 'row-layout',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [{ op: 'setRowHeight', sheet: 'Sheet1', row: 1, height: 24 }])
					return applyJournal(wb, [{ op: 'setRowHeight', sheet: 'Sheet1', row: 1, height: 32 }])
				},
			},
			{
				surface: 'column-layout',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 12 }])
					return applyJournal(wb, [{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 18 }])
				},
			},
			{
				surface: 'page-setup',
				exact: false,
				issue: { surface: 'page-setup', reason: 'page-setup-unsettable' },
				run: () => {
					const wb = AscendWorkbook.create()
					return applyJournal(wb, [
						{ op: 'setPageSetup', sheet: 'Sheet1', setup: { orientation: 'landscape' } },
					])
				},
			},
			{
				surface: 'sheet-layout',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [{ op: 'freezePane', sheet: 'Sheet1', row: 1, col: 1 }])
					return applyJournal(wb, [{ op: 'freezePane', sheet: 'Sheet1', row: 2, col: 0 }])
				},
			},
			{
				surface: 'x14-metadata',
				exact: false,
				issue: { surface: 'x14-metadata', reason: 'x14-metadata' },
				run: () => {
					const wb = AscendWorkbook.create()
					const sheet = wb.getWorkbookModel().getSheet('Sheet1')
					if (!sheet) throw new Error('missing sheet')
					sheet.x14DataValidations.push({
						index: 0,
						sqref: 'A1:A1',
						type: 'list',
						formula1: '"A,B"',
					})
					return applyJournal(wb, [
						{ op: 'copyRange', sheet: 'Sheet1', source: 'A1', target: 'D1', mode: 'validations' },
					])
				},
			},
			{
				surface: 'drawings',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					const sheet = wb.getWorkbookModel().getSheet('Sheet1')
					if (!sheet) throw new Error('missing sheet')
					sheet.drawingObjectRefs.push({
						drawingPartPath: 'xl/drawings/drawing1.xml',
						kind: 'textBox',
						id: 1,
						name: 'Note',
						text: 'before',
					})
					return applyJournal(wb, [
						{
							op: 'setDrawingText',
							sheet: 'Sheet1',
							drawingPartPath: 'xl/drawings/drawing1.xml',
							id: 1,
							text: 'after',
						},
					])
				},
			},
			{
				surface: 'charts',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					wb.getWorkbookModel().chartParts.push({
						partPath: 'xl/charts/chart1.xml',
						sheetName: 'Sheet1',
						series: [
							{
								nameRef: 'Sheet1!$B$1',
								categoryRef: 'Sheet1!$A$2:$A$4',
								valueRef: 'Sheet1!$B$2:$B$4',
							},
						],
					})
					return applyJournal(wb, [
						{
							op: 'setChartSeriesSource',
							partPath: 'xl/charts/chart1.xml',
							seriesIndex: 0,
							nameRef: 'Sheet1!$C$1',
							categoryRef: 'Sheet1!$A$2:$A$8',
							valueRef: 'Sheet1!$C$2:$C$8',
						},
					])
				},
			},
			{
				surface: 'pivot-caches',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					wb.getWorkbookModel().pivotTables.push({
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
					wb.getWorkbookModel().pivotCaches.push({
						partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
						cacheId: 1,
						sourceSheet: 'Raw',
						sourceRef: 'A1:B10',
						refreshOnLoad: false,
						enableRefresh: true,
						invalid: false,
						saveData: true,
						fields: [],
					})
					return applyJournal(wb, [
						{
							op: 'setPivotCache',
							pivotTable: 'PivotTable1',
							sourceSheet: 'Raw',
							sourceRef: 'A1:C20',
							refreshOnLoad: true,
							enableRefresh: false,
							invalid: true,
							saveData: false,
						},
					])
				},
			},
			{
				surface: 'workbook-metadata',
				exact: true,
				run: () => {
					const wb = AscendWorkbook.create()
					applyExact(wb, [
						{ op: 'setDocumentProperties', properties: { title: 'Before' }, mode: 'replace' },
					])
					return applyJournal(wb, [
						{ op: 'setDocumentProperties', properties: { title: 'After' }, mode: 'replace' },
					])
				},
			},
			{
				surface: 'package-parts',
				exact: false,
				issue: { surface: 'package-parts', reason: 'package-part-preservation' },
				run: () => {
					const wb = AscendWorkbook.create()
					return applyJournal(wb, [
						{
							op: 'setStyle',
							sheet: 'Sheet1',
							range: 'A1:A1',
							style: { font: { bold: true } },
						},
					])
				},
			},
		]

		const exercised = new Set<MutationJournalSurface>()
		for (const entry of cases) {
			exercised.add(entry.surface)
			const journal = entry.run()
			expect(journal.supported).toBe(true)
			expect(journal.exact).toBe(entry.exact)
			if (entry.issue) {
				const classified = journal.issues.map((issue) => classifyMutationJournalIssue(issue))
				expect(classified).toContainEqual({
					surface: entry.issue.surface,
					reason: entry.issue.reason,
					exactness: classifyMutationJournalSurface(entry.issue.surface).exactness,
					publicInverse: classifyMutationJournalSurface(entry.issue.surface).publicInverse,
				})
			} else {
				expect(journal.issues).toEqual([])
			}
		}

		for (const surface of [
			'cells',
			'formulas',
			'shared-formulas',
			'dynamic-arrays',
			'legacy-arrays',
			'data-tables',
			'spills',
			'tables',
			'defined-names',
			'comments',
			'hyperlinks',
			'data-validations',
			'conditional-formats',
			'auto-filters',
			'merged-cells',
			'row-layout',
			'column-layout',
			'page-setup',
			'sheet-layout',
			'x14-metadata',
			'drawings',
			'charts',
			'pivot-caches',
			'workbook-metadata',
			'package-parts',
		] satisfies readonly MutationJournalSurface[]) {
			expect(exercised.has(surface)).toBe(true)
		}
	})

	test('range operation journals mark imported formula-binding detachments lossy', () => {
		const cases: readonly {
			readonly name: string
			readonly setup: (workbook: AscendWorkbook) => void
			readonly ops: readonly Operation[]
			readonly issues: readonly {
				readonly ref: string
				readonly surface: MutationJournalSurface
			}[]
		}[] = [
			{
				name: 'copyRange target shared formula',
				setup: (wb) => {
					const sheet = seedSharedFormulaPair(wb, 0, 0)
					sheet.cells.set(0, 3, { value: numberValue(9), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'copyRange', sheet: 'Sheet1', source: 'D1', target: 'A2', mode: 'values' }],
				issues: [
					{ ref: 'Sheet1!A1', surface: 'shared-formulas' },
					{ ref: 'Sheet1!A2', surface: 'shared-formulas' },
				],
			},
			{
				name: 'copyRange target dynamic array',
				setup: (wb) => {
					const sheet = seedDynamicArrayFootprint(wb, 0)
					sheet.cells.set(0, 3, { value: numberValue(9), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'copyRange', sheet: 'Sheet1', source: 'D1', target: 'A2', mode: 'values' }],
				issues: [
					{ ref: 'Sheet1!A1', surface: 'dynamic-arrays' },
					{ ref: 'Sheet1!A2', surface: 'spills' },
					{ ref: 'Sheet1!A3', surface: 'spills' },
				],
			},
			{
				name: 'copyRange target data table',
				setup: (wb) => {
					const sheet = seedDataTableFormula(wb)
					sheet.cells.set(0, 5, { value: numberValue(9), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'copyRange', sheet: 'Sheet1', source: 'F1', target: 'C4', mode: 'values' }],
				issues: [{ ref: 'Sheet1!C3', surface: 'data-tables' }],
			},
			{
				name: 'copyRange target blocked spill',
				setup: (wb) => {
					const sheet = seedBlockedSpill(wb)
					sheet.cells.set(0, 3, { value: numberValue(9), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'copyRange', sheet: 'Sheet1', source: 'D1', target: 'A2', mode: 'values' }],
				issues: [{ ref: 'Sheet1!A1', surface: 'spills' }],
			},
			{
				name: 'moveRange target shared formula',
				setup: (wb) => {
					const sheet = seedSharedFormulaPair(wb, 0, 0)
					sheet.cells.set(0, 3, { value: numberValue(9), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'moveRange', sheet: 'Sheet1', source: 'D1', target: 'A2', mode: 'all' }],
				issues: [
					{ ref: 'Sheet1!A1', surface: 'shared-formulas' },
					{ ref: 'Sheet1!A2', surface: 'shared-formulas' },
				],
			},
			{
				name: 'moveRange target dynamic array',
				setup: (wb) => {
					const sheet = seedDynamicArrayFootprint(wb, 0)
					sheet.cells.set(0, 3, { value: numberValue(9), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'moveRange', sheet: 'Sheet1', source: 'D1', target: 'A2', mode: 'all' }],
				issues: [
					{ ref: 'Sheet1!A1', surface: 'dynamic-arrays' },
					{ ref: 'Sheet1!A2', surface: 'spills' },
					{ ref: 'Sheet1!A3', surface: 'spills' },
				],
			},
			{
				name: 'moveRange target data table',
				setup: (wb) => {
					const sheet = seedDataTableFormula(wb)
					sheet.cells.set(0, 5, { value: numberValue(9), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'moveRange', sheet: 'Sheet1', source: 'F1', target: 'C4', mode: 'all' }],
				issues: [{ ref: 'Sheet1!C3', surface: 'data-tables' }],
			},
			{
				name: 'moveRange target blocked spill',
				setup: (wb) => {
					const sheet = seedBlockedSpill(wb)
					sheet.cells.set(0, 3, { value: numberValue(9), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'moveRange', sheet: 'Sheet1', source: 'D1', target: 'A2', mode: 'all' }],
				issues: [{ ref: 'Sheet1!A1', surface: 'spills' }],
			},
			{
				name: 'sortRange shared formula',
				setup: (wb) => {
					const sheet = seedSharedFormulaPair(wb, 0, 3)
					sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: DEFAULT_STYLE_ID })
					sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'sortRange', sheet: 'Sheet1', range: 'A1:D2', by: [{ column: 'A' }] }],
				issues: [
					{ ref: 'Sheet1!D1', surface: 'shared-formulas' },
					{ ref: 'Sheet1!D2', surface: 'shared-formulas' },
				],
			},
			{
				name: 'sortRange dynamic array',
				setup: (wb) => {
					const sheet = seedDynamicArrayFootprint(wb, 3)
					sheet.cells.set(0, 0, { value: numberValue(1), formula: null, styleId: DEFAULT_STYLE_ID })
					sheet.cells.set(1, 0, { value: numberValue(2), formula: null, styleId: DEFAULT_STYLE_ID })
					sheet.cells.set(2, 0, { value: numberValue(3), formula: null, styleId: DEFAULT_STYLE_ID })
				},
				ops: [{ op: 'sortRange', sheet: 'Sheet1', range: 'A1:D3', by: [{ column: 'A' }] }],
				issues: [
					{ ref: 'Sheet1!D1', surface: 'dynamic-arrays' },
					{ ref: 'Sheet1!D2', surface: 'spills' },
					{ ref: 'Sheet1!D3', surface: 'spills' },
				],
			},
			{
				name: 'sortRange data table',
				setup: seedDataTableFormula,
				ops: [{ op: 'sortRange', sheet: 'Sheet1', range: 'A1:C5', by: [{ column: 'A' }] }],
				issues: [{ ref: 'Sheet1!C3', surface: 'data-tables' }],
			},
			{
				name: 'sortRange blocked spill',
				setup: seedBlockedSpill,
				ops: [{ op: 'sortRange', sheet: 'Sheet1', range: 'A1:A2', by: [{ column: 'A' }] }],
				issues: [{ ref: 'Sheet1!A1', surface: 'spills' }],
			},
		]

		for (const entry of cases) {
			const wb = AscendWorkbook.create()
			entry.setup(wb)

			const journal = applyJournal(wb, entry.ops)

			expect(journal.supported, entry.name).toBe(true)
			expect(journal.exact, entry.name).toBe(false)
			for (const issue of entry.issues) {
				expect(journal.issues, `${entry.name} ${issue.ref}`).toContainEqual(
					expect.objectContaining({
						code: 'LOSSY_INVERSE',
						surface: issue.surface,
						reason: 'formula-binding-metadata',
						refs: [issue.ref],
					}),
				)
			}
		}
	})

	test('destructive cell journals mark formula-binding detachments lossy', () => {
		const cases: readonly {
			readonly name: string
			readonly ops: readonly Operation[]
		}[] = [
			{
				name: 'literal write',
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A2', value: 7 }] }],
			},
			{
				name: 'formula write',
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'A2', formula: 'B2*3' }],
			},
			{
				name: 'fill formula',
				ops: [{ op: 'fillFormula', sheet: 'Sheet1', range: 'A2', formula: 'B2*3' }],
			},
			{
				name: 'formula clear',
				ops: [{ op: 'clearRange', sheet: 'Sheet1', range: 'A2', what: 'formulas' }],
			},
			{
				name: 'full clear',
				ops: [{ op: 'clearRange', sheet: 'Sheet1', range: 'A2', what: 'all' }],
			},
			{
				name: 'rich text write',
				ops: [{ op: 'setRichText', sheet: 'Sheet1', ref: 'A2', runs: [{ text: 'manual' }] }],
			},
		]

		for (const entry of cases) {
			const wb = AscendWorkbook.create()
			seedSharedFormulaPair(wb, 0, 0)

			const changed = wb.apply(entry.ops, { journal: true })

			expect(changed.errors, entry.name).toEqual([])
			expect(
				changed.affectedCells.map((ref) => ref.replace(/^Sheet1!/, '')),
				entry.name,
			).toEqual(['A1', 'A2'])
			expect(changed.journal?.supported, entry.name).toBe(true)
			expect(changed.journal?.exact, entry.name).toBe(false)
			expect(changed.journal, entry.name).toBeDefined()
			if (!changed.journal) throw new Error(`missing journal for ${entry.name}`)
			expect(journalIssueRefs(changed.journal, 'shared-formulas')).toEqual([
				'Sheet1!A1',
				'Sheet1!A2',
			])
			expect(
				(changed.journal?.issues ?? [])
					.filter((issue) => issue.surface === 'shared-formulas')
					.every((issue) => issue.reason === 'formula-binding-metadata'),
				entry.name,
			).toBe(true)
		}
	})

	test('style-only journals preserve formula bindings without binding loss issues', () => {
		const cases: readonly {
			readonly name: string
			readonly ops: readonly Operation[]
			readonly exact: boolean
		}[] = [
			{
				name: 'setStyle',
				ops: [
					{
						op: 'setStyle',
						sheet: 'Sheet1',
						range: 'A2:A2',
						style: { font: { bold: true } },
					},
				],
				exact: false,
			},
			{
				name: 'style clear',
				ops: [{ op: 'clearRange', sheet: 'Sheet1', range: 'A2', what: 'styles' }],
				exact: true,
			},
			{
				name: 'format copy',
				ops: [{ op: 'copyRange', sheet: 'Sheet1', source: 'B1', target: 'A2', mode: 'formats' }],
				exact: true,
			},
		]

		for (const entry of cases) {
			const wb = AscendWorkbook.create()
			seedSharedFormulaPair(wb, 0, 0)

			const changed = wb.apply(entry.ops, { journal: true })

			expect(changed.errors, entry.name).toEqual([])
			expect(changed.journal?.supported, entry.name).toBe(true)
			expect(changed.journal?.exact, entry.name).toBe(entry.exact)
			expect(changed.journal, entry.name).toBeDefined()
			if (!changed.journal) throw new Error(`missing journal for ${entry.name}`)
			expect(journalIssueRefs(changed.journal, 'shared-formulas'), entry.name).toEqual([])
			expect(formulaInfoRefsIn(wb, 'Sheet1', ['A1', 'A2']), entry.name).toEqual([
				'Sheet1!A1',
				'Sheet1!A2',
			])
		}
	})

	test('real XLSX shared formula member edits journal every detached imported peer', async () => {
		const wb = await AscendWorkbook.open(
			readFileSync(new URL('../../../fixtures/xlsx/poi/shared_formulas.xlsx', import.meta.url)),
		)
		const sharedRefs = Array.from({ length: 40 }, (_, index) => `Label!A${index + 2}`)
		expect(collectFormulaInfoRefs(wb, 'Label')).toEqual(sharedRefs)

		const changed = wb.apply(
			[{ op: 'setCells', sheet: 'Label', updates: [{ ref: 'A3', value: 99 }] }],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.affectedCells).toEqual(sharedRefs.map((ref) => ref.split('!')[1]))
		expect(collectFormulaInfoRefs(wb, 'Label')).toEqual([])
		expect(changed.journal).toBeDefined()
		if (!changed.journal) throw new Error('missing journal')
		expect(changed.journal.exact).toBe(false)
		expect(journalIssueRefs(changed.journal, 'shared-formulas')).toEqual([...sharedRefs].sort())
		expect(
			changed.journal.issues
				.filter((issue) => issue.surface === 'shared-formulas')
				.every((issue) => issue.reason === 'formula-binding-metadata'),
		).toBe(true)
	})

	test('real XLSX data table member edits journal detached table metadata', async () => {
		const wb = await AscendWorkbook.open(
			readFileSync(
				new URL(
					'../../../fixtures/xlsx/closedxml/Other_Formulas_DataTableFormula-Excel-Input.xlsx',
					import.meta.url,
				),
			),
		)
		expect(collectFormulaInfoRefs(wb, '1D Row')).toEqual(['1D Row!C4'])

		const changed = wb.apply(
			[{ op: 'setCells', sheet: '1D Row', updates: [{ ref: 'C6', value: 99 }] }],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect([...changed.affectedCells].sort()).toEqual(['C4', 'C6'])
		expect(collectFormulaInfoRefs(wb, '1D Row')).toEqual([])
		expect(changed.journal).toBeDefined()
		if (!changed.journal) throw new Error('missing journal')
		expect(changed.journal.exact).toBe(false)
		expect(journalIssueRefs(changed.journal, 'data-tables')).toEqual(['1D Row!C4'])
		expect(
			changed.journal.issues
				.filter((issue) => issue.surface === 'data-tables')
				.every((issue) => issue.reason === 'formula-binding-metadata'),
		).toBe(true)

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(collectFormulaInfoRefs(reopened, '1D Row')).toEqual([])
		expect(cellFormulas(reopened, '1D Row', ['C4', 'C6'])).toEqual({ C4: null, C6: null })
	})

	test('real XLSX legacy array destructive edits fail without dropping metadata', async () => {
		const source = readFileSync(
			new URL('../../../fixtures/xlsx/closedxml/Other_Formulas_ArrayFormula.xlsx', import.meta.url),
		)
		const wb = await AscendWorkbook.open(source)
		expect(collectFormulaInfoRefs(wb, 'Sheet1')).toEqual(['Sheet1!A1'])

		const changed = wb.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B2', value: 99 }] }],
			{ journal: true },
		)

		expect(changed.errors).toEqual([
			expect.objectContaining({
				code: 'VALIDATION_ERROR',
				refs: ['B2', 'A1:B2'],
			}),
		])
		expect(changed.affectedCells).toEqual([])
		expect(changed.journal).toEqual(
			expect.objectContaining({
				supported: false,
				exact: false,
				issues: [
					expect.objectContaining({
						code: 'JOURNAL_UNAVAILABLE',
						surface: 'package-parts',
						reason: 'journal-unavailable',
					}),
				],
			}),
		)
		expect(collectFormulaInfoRefs(wb, 'Sheet1')).toEqual(['Sheet1!A1'])
		expect(cellFormulas(wb, 'Sheet1', ['A1', 'B2'])).toEqual({ A1: '1+2', B2: undefined })

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(collectFormulaInfoRefs(reopened, 'Sheet1')).toEqual(['Sheet1!A1'])
		expect(cellFormulas(reopened, 'Sheet1', ['A1', 'B2'])).toEqual({ A1: '1+2', B2: undefined })
	})

	test('real XLSX legacy array style edits preserve bindings through save reopen', async () => {
		const wb = await AscendWorkbook.open(
			readFileSync(
				new URL(
					'../../../fixtures/xlsx/closedxml/Other_Formulas_ArrayFormula.xlsx',
					import.meta.url,
				),
			),
		)
		expect(wb.check().valid).toBe(true)
		expect(collectFormulaInfoRefs(wb, 'Sheet1')).toEqual(['Sheet1!A1'])

		const changed = wb.apply(
			[
				{
					op: 'setStyle',
					sheet: 'Sheet1',
					range: 'A1:A1',
					style: { numberFormat: '0.0' },
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.affectedCells).toEqual(['A1'])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(false)
		expect(changed.journal?.issues).toEqual([
			expect.objectContaining({
				surface: 'package-parts',
				reason: 'package-part-preservation',
				refs: ['Sheet1!A1'],
			}),
		])
		expect(journalIssueRefs(changed.journal, 'legacy-arrays')).toEqual([])
		expect(wb.check().valid).toBe(true)
		expect(collectFormulaInfoRefs(wb, 'Sheet1')).toEqual(['Sheet1!A1'])
		expect(cellFormulas(wb, 'Sheet1', ['A1'])).toEqual({ A1: '1+2' })

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(reopened.check().valid).toBe(true)
		expect(collectFormulaInfoRefs(reopened, 'Sheet1')).toEqual(['Sheet1!A1'])
		expect(cellFormulas(reopened, 'Sheet1', ['A1'])).toEqual({ A1: '1+2' })
		expect(reopened.sheet('Sheet1')?.cell('A1')?.formulaBinding).toEqual({
			kind: 'array',
			ref: 'A1:B2',
		})
	})

	test('real XLSX multi-axis shared formula edits materialize sibling formulas', async () => {
		const wb = await AscendWorkbook.open(
			readFileSync(
				new URL(
					'../../../fixtures/xlsx/calamine/issue_565_multi_axis_shared.xlsx',
					import.meta.url,
				),
			),
		)
		const groupRefs = ['B1', 'C1', 'D1', 'B2', 'C2', 'D2']
		expect(formulaInfoRefsIn(wb, 'Sheet1', groupRefs)).toEqual(
			groupRefs.map((ref) => `Sheet1!${ref}`),
		)

		const changed = wb.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'C2', value: 99 }] }],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.affectedCells).toEqual(groupRefs)
		expect(formulaInfoRefsIn(wb, 'Sheet1', groupRefs)).toEqual([])
		expect(cellFormulas(wb, 'Sheet1', groupRefs)).toEqual({
			B1: 'B1',
			C1: 'C1',
			D1: 'D1',
			B2: 'B2',
			C2: null,
			D2: 'D2',
		})
		expect(changed.journal).toBeDefined()
		if (!changed.journal) throw new Error('missing journal')
		expect(journalIssueRefs(changed.journal, 'shared-formulas')).toEqual(
			groupRefs.map((ref) => `Sheet1!${ref}`).sort(),
		)
	})

	test('copySheet exact journals preserve retargeted formula binding evidence', () => {
		const wb = AscendWorkbook.create()
		const source = wb.getWorkbookModel().getSheet('Sheet1')
		if (!source) throw new Error('missing Sheet1')
		source.cells.set(0, 0, {
			value: numberValue(1),
			formula: 'B1*2',
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'journal-copy-shared',
				isMaster: true,
				masterRef: 'Sheet1!A1',
				ref: 'Sheet1!A1:A2',
			},
		})
		source.cells.set(1, 0, {
			value: numberValue(2),
			formula: null,
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'journal-copy-shared',
				isMaster: false,
				masterRef: 'Sheet1!A1',
				ref: 'Sheet1!A1:A2',
			},
		})
		source.cells.set(3, 0, {
			value: numberValue(4),
			formula: 'Sheet1!A4:Sheet1!B5',
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: { kind: 'array', ref: 'Sheet1!A4:B5' },
		})
		source.cells.set(6, 0, {
			value: numberValue(7),
			formula: null,
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: {
				kind: 'dataTable',
				ref: 'Sheet1!A7:B8',
				dtr: true,
				r1: 'Sheet1!C1',
				r2: 'Sheet1!D1',
			},
		})
		const before = journalEvidence(wb)

		const changed = wb.apply([{ op: 'copySheet', sheet: 'Sheet1', newName: 'Copy' }], {
			journal: true,
		})

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.issues).toEqual([])
		const copy = wb.getWorkbookModel().getSheet('Copy')
		expect(copy?.cells.get(0, 0)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: 'journal-copy-shared',
			isMaster: true,
			masterRef: 'Copy!A1',
			ref: 'Copy!A1:A2',
		})
		expect(copy?.cells.get(1, 0)?.formulaInfo).toEqual({
			kind: 'shared',
			sharedIndex: 'journal-copy-shared',
			isMaster: false,
			masterRef: 'Copy!A1',
			ref: 'Copy!A1:A2',
		})
		expect(copy?.cells.get(3, 0)?.formulaInfo).toEqual({ kind: 'array', ref: 'Copy!A4:B5' })
		expect(copy?.cells.get(6, 0)?.formulaInfo).toEqual({
			kind: 'dataTable',
			ref: 'Copy!A7:B8',
			dtr: true,
			r1: 'Copy!C1',
			r2: 'Copy!D1',
		})

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalEvidence(wb)).toEqual(before)
	})

	test('copySheet saved workbooks reopen with formula text and binding metadata retargeted together', async () => {
		const wb = AscendWorkbook.create()
		const source = wb.getWorkbookModel().getSheet('Sheet1')
		if (!source) throw new Error('missing Sheet1')
		source.cells.set(0, 0, {
			value: numberValue(2),
			formula: 'Sheet1!B1*2',
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'journal-copy-shared',
				isMaster: true,
				masterRef: 'Sheet1!A1',
				ref: 'Sheet1!A1:A2',
			},
		})
		source.cells.set(1, 0, {
			value: numberValue(4),
			formula: null,
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'journal-copy-shared',
				isMaster: false,
				masterRef: 'Sheet1!A1',
				ref: 'Sheet1!A1:A2',
			},
		})
		source.cells.set(3, 0, {
			value: numberValue(4),
			formula: 'SUM(Sheet1!C4:Sheet1!D5)',
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: { kind: 'array', ref: 'Sheet1!A4:B5' },
		})
		source.cells.set(6, 0, {
			value: numberValue(7),
			formula: null,
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: {
				kind: 'dataTable',
				ref: 'Sheet1!A7:B8',
				dtr: true,
				r1: 'Sheet1!C1',
				r2: 'Sheet1!D1',
			},
		})

		const changed = wb.apply([{ op: 'copySheet', sheet: 'Sheet1', newName: 'Copy' }], {
			journal: true,
		})
		expect(changed.errors).toEqual([])
		expect(wb.check().valid).toBe(true)

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(reopened.check().valid).toBe(true)
		expect(reopened.formula('Copy!A1')?.normalizedFormula).toBe('Copy!B1*2')
		expect(reopened.formula('Copy!A2')?.normalizedFormula).toBe('Copy!B2*2')
		expect(reopened.formula('Copy!A4')?.normalizedFormula).toBe('SUM(Copy!C4:D5)')
		expect(reopened.sheet('Copy')?.cell('A1')?.formulaBinding).toEqual({
			kind: 'shared',
			sharedIndex: 'journal-copy-shared',
			isMaster: true,
			masterRef: 'A1',
			ref: 'Copy!A1:A2',
		})
		expect(reopened.sheet('Copy')?.cell('A2')?.formulaBinding).toEqual({
			kind: 'shared',
			sharedIndex: 'journal-copy-shared',
			isMaster: false,
			masterRef: 'A1',
		})
		expect(reopened.sheet('Copy')?.cell('A4')?.formulaBinding).toEqual({
			kind: 'array',
			ref: 'Copy!A4:B5',
		})
		expect(reopened.sheet('Copy')?.cell('A7')?.formulaBinding).toEqual({
			kind: 'dataTable',
			ref: 'Copy!A7:B8',
			dtr: true,
			r1: 'Copy!C1',
			r2: 'Copy!D1',
		})
	})

	test('copySheet retargets imported sheet-qualified shared formula text and bindings across save reopen', async () => {
		const source = AscendWorkbook.create()
		const sheet = source.getWorkbookModel().getSheet('Sheet1')
		if (!sheet) throw new Error('missing Sheet1')
		sheet.cells.set(0, 0, {
			value: numberValue(2),
			formula: 'Sheet1!B1*2',
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'imported-copy-shared',
				isMaster: true,
				masterRef: 'Sheet1!A1',
				ref: 'Sheet1!A1:A2',
			},
		})
		sheet.cells.set(1, 0, {
			value: numberValue(4),
			formula: null,
			styleId: DEFAULT_STYLE_ID,
			formulaInfo: {
				kind: 'shared',
				sharedIndex: 'imported-copy-shared',
				isMaster: false,
				masterRef: 'Sheet1!A1',
				ref: 'Sheet1!A1:A2',
			},
		})

		const imported = await AscendWorkbook.open(source.toBytes())
		expect(imported.check().valid).toBe(true)
		expect(imported.formula('Sheet1!A1')?.normalizedFormula).toBe('Sheet1!B1*2')
		expect(imported.formula('Sheet1!A2')?.normalizedFormula).toBe('Sheet1!B2*2')

		const changed = imported.apply([{ op: 'copySheet', sheet: 'Sheet1', newName: 'Copy' }], {
			journal: true,
		})

		expect(changed.errors).toEqual([])
		expect(imported.check().valid).toBe(true)
		const reopened = await AscendWorkbook.open(imported.toBytes())
		expect(reopened.check().valid).toBe(true)
		expect(reopened.formula('Copy!A1')?.normalizedFormula).toBe('Copy!B1*2')
		expect(reopened.formula('Copy!A2')?.normalizedFormula).toBe('Copy!B2*2')
		expect(reopened.sheet('Copy')?.cell('A1')?.formulaBinding).toEqual({
			kind: 'shared',
			sharedIndex: 'imported-copy-shared',
			isMaster: true,
			masterRef: 'A1',
			ref: 'Copy!A1:A2',
		})
		expect(reopened.sheet('Copy')?.cell('A2')?.formulaBinding).toEqual({
			kind: 'shared',
			sharedIndex: 'imported-copy-shared',
			isMaster: false,
			masterRef: 'A1',
		})
	})

	test('copySheet preserves imported shared formula text and bindings across save reopen', async () => {
		const wb = await AscendWorkbook.open(
			readFileSync(new URL('../../../fixtures/xlsx/poi/shared_formulas.xlsx', import.meta.url)),
		)
		expect(wb.check().valid).toBe(true)
		expect(collectFormulaInfoRefs(wb, 'Label')).toEqual(
			Array.from({ length: 40 }, (_, index) => `Label!A${index + 2}`),
		)

		const changed = wb.apply([{ op: 'copySheet', sheet: 'Label', newName: 'Label_Copy' }], {
			journal: true,
		})

		expect(changed.errors).toEqual([])
		expect(changed.journal).toEqual(
			expect.objectContaining({
				supported: true,
				exact: false,
				issues: [
					expect.objectContaining({
						surface: 'package-parts',
						reason: 'package-part-preservation',
						refs: ['sheet:Label', 'sheet:Label_Copy'],
					}),
				],
			}),
		)
		expect(wb.check().valid).toBe(true)

		const reopened = await AscendWorkbook.open(wb.toBytes())
		expect(reopened.check().valid).toBe(true)
		expect(collectFormulaInfoRefs(reopened, 'Label_Copy')).toEqual(
			Array.from({ length: 40 }, (_, index) => `Label_Copy!A${index + 2}`),
		)
		expect(reopened.formula('Label_Copy!A2')?.normalizedFormula).toBe('B2')
		expect(reopened.formula('Label_Copy!A3')?.normalizedFormula).toBe('B3')
		expect(reopened.sheet('Label_Copy')?.cell('A2')?.formulaBinding).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: true,
			masterRef: 'A2',
			ref: 'A2:A41',
		})
		expect(reopened.sheet('Label_Copy')?.cell('A3')?.formulaBinding).toEqual({
			kind: 'shared',
			sharedIndex: '0',
			isMaster: false,
			masterRef: 'A2',
		})
	})

	test('representative exact journals restore full workbook evidence after inverse apply', () => {
		const cases: readonly {
			readonly name: string
			readonly setup: (workbook: AscendWorkbook) => void
			readonly ops: readonly Operation[]
		}[] = [
			{
				name: 'formula cache replacement',
				setup: (wb) => {
					applyExact(wb, [
						{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
						{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
					])
					const recalc = wb.recalc()
					expect(recalc.errors).toEqual([])
				},
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*3' }],
			},
			{
				name: 'copyRange target formula cache',
				setup: (wb) => {
					applyExact(wb, [
						{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'source' }] },
						{ op: 'setFormula', sheet: 'Sheet1', ref: 'D1', formula: '2+3' },
					])
					const recalc = wb.recalc()
					expect(recalc.errors).toEqual([])
				},
				ops: [{ op: 'copyRange', sheet: 'Sheet1', source: 'A1', target: 'D1', mode: 'all' }],
			},
			{
				name: 'moveRange source and target formula caches',
				setup: (wb) => {
					applyExact(wb, [
						{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '2+3' },
						{ op: 'setFormula', sheet: 'Sheet1', ref: 'D1', formula: '9+1' },
					])
					const recalc = wb.recalc()
					expect(recalc.errors).toEqual([])
				},
				ops: [{ op: 'moveRange', sheet: 'Sheet1', source: 'A1', target: 'D1', mode: 'all' }],
			},
			{
				name: 'sortRange formula caches',
				setup: (wb) => {
					applyExact(wb, [
						{
							op: 'setCells',
							sheet: 'Sheet1',
							updates: [
								{ ref: 'A1', value: 'Key' },
								{ ref: 'B1', value: 'Score' },
								{ ref: 'A2', value: 'b' },
								{ ref: 'A3', value: 'a' },
							],
						},
						{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: '2+1' },
						{ op: 'setFormula', sheet: 'Sheet1', ref: 'B3', formula: '1+1' },
					])
					const recalc = wb.recalc()
					expect(recalc.errors).toEqual([])
				},
				ops: [{ op: 'sortRange', sheet: 'Sheet1', range: 'A1:B3', by: [{ column: 'Score' }] }],
			},
			{
				name: 'copyRange all metadata',
				setup: (wb) => {
					applyExact(wb, [
						{
							op: 'setCells',
							sheet: 'Sheet1',
							updates: [
								{ ref: 'A1', value: 'source' },
								{ ref: 'D1', value: 'target' },
							],
						},
						{ op: 'setStyle', sheet: 'Sheet1', range: 'A1:A1', style: { numberFormat: '0.0' } },
						{ op: 'setComment', sheet: 'Sheet1', ref: 'A1', text: 'copy me', author: 'agent' },
						{ op: 'setComment', sheet: 'Sheet1', ref: 'D1', text: 'keep me', author: 'analyst' },
						{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', url: 'https://source.example' },
						{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'D1', url: 'https://target.example' },
						{
							op: 'setDataValidation',
							sheet: 'Sheet1',
							range: 'A1:A1',
							rule: { type: 'list', formula1: '"A,B"' },
						},
						{
							op: 'setConditionalFormat',
							sheet: 'Sheet1',
							range: 'A1:A1',
							rule: { type: 'expression', formula: 'A1<>""', priority: 1 },
						},
					])
				},
				ops: [{ op: 'copyRange', sheet: 'Sheet1', source: 'A1:A1', target: 'D1', mode: 'all' }],
			},
			{
				name: 'moveRange formula surface rewrites',
				setup: (wb) => {
					applyExact(wb, [
						{
							op: 'setCells',
							sheet: 'Sheet1',
							updates: [
								{ ref: 'A1', value: 1 },
								{ ref: 'C1', value: 3 },
							],
						},
						{ op: 'setFormula', sheet: 'Sheet1', ref: 'D1', formula: 'A1+C1' },
						{ op: 'setDefinedName', name: 'Input', ref: 'Sheet1!$A$1' },
						{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'B1', location: 'Sheet1!A1' },
					])
				},
				ops: [{ op: 'moveRange', sheet: 'Sheet1', source: 'A1:A1', target: 'C1', mode: 'all' }],
			},
			{
				name: 'sortRange row metadata',
				setup: (wb) => {
					applyExact(wb, [
						{
							op: 'setCells',
							sheet: 'Sheet1',
							updates: [
								{ ref: 'A1', value: 'Key' },
								{ ref: 'B1', value: 'Qty' },
								{ ref: 'A2', value: 'b' },
								{ ref: 'B2', value: 2 },
								{ ref: 'A3', value: 'a' },
								{ ref: 'B3', value: 1 },
							],
						},
						{
							op: 'setDataValidation',
							sheet: 'Sheet1',
							range: 'A2:A3',
							rule: { type: 'list', formula1: '"a,b"' },
						},
						{
							op: 'setConditionalFormat',
							sheet: 'Sheet1',
							range: 'B2:B3',
							rule: { type: 'cellIs', operator: 'greaterThan', formula: '0', priority: 1 },
						},
						{ op: 'setComment', sheet: 'Sheet1', ref: 'A2', text: 'row note', author: 'agent' },
						{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'B2', url: 'https://row.example' },
					])
				},
				ops: [{ op: 'sortRange', sheet: 'Sheet1', range: 'A1:B3', by: [{ column: 'Qty' }] }],
			},
			{
				name: 'page setup and print area',
				setup: (wb) => {
					applyExact(wb, [
						{
							op: 'setPageSetup',
							sheet: 'Sheet1',
							setup: { orientation: 'portrait', paperSize: 9 },
						},
						{ op: 'setPrintArea', sheet: 'Sheet1', range: 'A1:B2' },
					])
				},
				ops: [
					{
						op: 'setPageSetup',
						sheet: 'Sheet1',
						setup: { orientation: 'landscape', paperSize: 9 },
					},
					{ op: 'setPrintArea', sheet: 'Sheet1', range: 'C1:D2' },
				],
			},
			{
				name: 'pre-existing style registry entry',
				setup: (wb) => {
					wb.getWorkbookModel().styles.register({ font: { bold: true }, numberFormat: '0.0' })
				},
				ops: [
					{
						op: 'setStyle',
						sheet: 'Sheet1',
						range: 'E9:E9',
						style: { font: { bold: true }, numberFormat: '0.0' },
					},
				],
			},
		]

		for (const entry of cases) {
			const wb = AscendWorkbook.create()
			entry.setup(wb)
			const before = journalEvidence(wb)
			const changed = wb.apply(entry.ops, { journal: true })
			expect(changed.errors, entry.name).toEqual([])
			expect(changed.journal?.supported, entry.name).toBe(true)
			expect(changed.journal?.exact, entry.name).toBe(true)
			expect(changed.journal?.issues, entry.name).toEqual([])

			const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
			expect(undo.errors, entry.name).toEqual([])
			expect(journalEvidence(wb), entry.name).toEqual(before)
		}
	})

	test('representative exact saved journals restore package bytes after inverse apply', async () => {
		const cases: readonly {
			readonly name: string
			readonly seedOps: readonly Operation[]
			readonly ops: readonly Operation[]
		}[] = [
			{
				name: 'hyperlink relationship',
				seedOps: [
					{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'link' }] },
					{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', url: 'https://example.com' },
				],
				ops: [{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', url: 'https://changed.example' }],
			},
			{
				name: 'theme metadata',
				seedOps: [
					{
						op: 'setTheme',
						themeName: 'Office',
						colorSchemeName: 'Office Colors',
						majorFontLatin: 'Aptos Display',
						minorFontLatin: 'Aptos',
						themeColors: [
							{ slot: 'accent1', rgb: '4F81BD' },
							{ slot: 'lt1', systemColor: 'window', lastColor: 'FFFFFF' },
						],
					},
				],
				ops: [
					{
						op: 'setTheme',
						themeName: 'Brand',
						colorSchemeName: 'Brand Colors',
						majorFontLatin: 'Inter Display',
						minorFontLatin: 'Inter',
						themeColors: [
							{ slot: 'accent1', rgb: '0F6CBD' },
							{ slot: 'lt1', systemColor: 'windowText', lastColor: '000000' },
						],
					},
				],
			},
			{
				name: 'print area addition',
				seedOps: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'print' }] }],
				ops: [{ op: 'setPrintArea', sheet: 'Sheet1', range: 'A1:A1' }],
			},
			{
				name: 'print area replacement',
				seedOps: [
					{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'print' }] },
					{ op: 'setPrintArea', sheet: 'Sheet1', range: 'A1:A1' },
				],
				ops: [{ op: 'setPrintArea', sheet: 'Sheet1', range: 'B1:B1' }],
			},
			{
				name: 'page setup replacement',
				seedOps: [
					{
						op: 'setPageSetup',
						sheet: 'Sheet1',
						setup: {
							orientation: 'portrait',
							paperSize: 1,
							scale: 95,
							margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75 },
						},
					},
				],
				ops: [
					{
						op: 'setPageSetup',
						sheet: 'Sheet1',
						setup: {
							orientation: 'landscape',
							margins: { left: 0.25, right: 0.25 },
						},
					},
				],
			},
			{
				name: 'tab color replacement',
				seedOps: [{ op: 'setTabColor', sheet: 'Sheet1', color: 'FF0000' }],
				ops: [{ op: 'setTabColor', sheet: 'Sheet1', color: '00FF00' }],
			},
			{
				name: 'sheet protection replacement',
				seedOps: [
					{
						op: 'setSheetProtection',
						sheet: 'Sheet1',
						password: 'before',
						options: { selectLockedCells: false },
					},
				],
				ops: [
					{
						op: 'setSheetProtection',
						sheet: 'Sheet1',
						password: 'after',
						options: { selectUnlockedCells: false },
					},
				],
			},
			{
				name: 'freeze pane replacement',
				seedOps: [{ op: 'freezePane', sheet: 'Sheet1', row: 1, col: 1 }],
				ops: [{ op: 'freezePane', sheet: 'Sheet1', row: 2, col: 0 }],
			},
			{
				name: 'existing row and column layout',
				seedOps: [
					{ op: 'setRowHeight', sheet: 'Sheet1', row: 1, height: 20 },
					{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 12 },
					{ op: 'hideRows', sheet: 'Sheet1', at: 1, count: 1, hidden: true },
					{ op: 'hideCols', sheet: 'Sheet1', at: 1, count: 1, hidden: true },
				],
				ops: [
					{ op: 'setRowHeight', sheet: 'Sheet1', row: 1, height: 25 },
					{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 18 },
					{ op: 'hideRows', sheet: 'Sheet1', at: 1, count: 1, hidden: false },
					{ op: 'hideCols', sheet: 'Sheet1', at: 1, count: 1, hidden: false },
				],
			},
			{
				name: 'data validation tail replacement and deletion',
				seedOps: [
					{
						op: 'setDataValidation',
						sheet: 'Sheet1',
						range: 'A1:A2',
						rule: {
							type: 'whole',
							operator: 'between',
							formula1: '1',
							formula2: '9',
							allowBlank: true,
							showErrorMessage: true,
						},
					},
					{
						op: 'setDataValidation',
						sheet: 'Sheet1',
						range: 'B1:B2',
						rule: { type: 'list', formula1: '"A,B"', allowBlank: true, showErrorMessage: true },
					},
				],
				ops: [
					{
						op: 'setDataValidation',
						sheet: 'Sheet1',
						range: 'B1:B2',
						rule: {
							type: 'whole',
							operator: 'greaterThan',
							formula1: '3',
							allowBlank: true,
							showErrorMessage: true,
						},
					},
					{ op: 'deleteDataValidation', sheet: 'Sheet1', range: 'B1:B2' },
				],
			},
			{
				name: 'data validation non-tail replacement',
				seedOps: [
					{
						op: 'setDataValidation',
						sheet: 'Sheet1',
						range: 'A1:A1',
						rule: { type: 'whole', formula1: '1', allowBlank: true, showErrorMessage: true },
					},
					{
						op: 'setDataValidation',
						sheet: 'Sheet1',
						range: 'B1:B1',
						rule: { type: 'whole', formula1: '2', allowBlank: true, showErrorMessage: true },
					},
				],
				ops: [
					{
						op: 'setDataValidation',
						sheet: 'Sheet1',
						range: 'A1:A1',
						rule: { type: 'whole', formula1: '3', allowBlank: true, showErrorMessage: true },
					},
				],
			},
			{
				name: 'conditional format tail replacement',
				seedOps: [
					{
						op: 'setConditionalFormat',
						sheet: 'Sheet1',
						range: 'A1:A2',
						rule: { type: 'expression', formula: 'A1>0', priority: 1 },
					},
				],
				ops: [
					{
						op: 'setConditionalFormat',
						sheet: 'Sheet1',
						range: 'A1:A2',
						rule: { type: 'expression', formula: 'A1>5', priority: 1 },
					},
				],
			},
			{
				name: 'auto filter criteria and clear',
				seedOps: [
					{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:B10' },
					{
						op: 'setAutoFilter',
						sheet: 'Sheet1',
						range: 'A1:B10',
						column: 0,
						values: ['Open'],
					},
				],
				ops: [
					{
						op: 'setAutoFilter',
						sheet: 'Sheet1',
						range: 'A1:B10',
						column: 0,
						values: ['Closed'],
					},
					{ op: 'clearAutoFilter', sheet: 'Sheet1' },
				],
			},
			{
				name: 'workbook view replacement',
				seedOps: [
					{
						op: 'setWorkbookView',
						view: { activeTab: 0, firstSheet: 0 },
						index: 0,
						mode: 'replace',
					},
				],
				ops: [
					{
						op: 'setWorkbookView',
						view: { activeTab: 0, firstSheet: 0, showSheetTabs: false },
						index: 0,
						mode: 'replace',
					},
				],
			},
			{
				name: 'calc settings replacement',
				seedOps: [{ op: 'setCalcSettings', settings: { mode: 'auto' } }],
				ops: [{ op: 'setCalcSettings', settings: { mode: 'manual', fullCalcOnLoad: true } }],
			},
			{
				name: 'semantic no-op cell edit',
				seedOps: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
			},
		]

		for (const entry of cases) {
			const seeded = AscendWorkbook.create()
			applyExact(seeded, entry.seedOps)
			const beforeBytes = seeded.toBytes()
			const wb = await AscendWorkbook.open(beforeBytes)
			const changed = wb.apply(entry.ops, { journal: true })
			expect(changed.errors, entry.name).toEqual([])
			expect(changed.journal?.supported, entry.name).toBe(true)
			expect(changed.journal?.exact, entry.name).toBe(true)

			const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
			expect(undo.errors, entry.name).toEqual([])
			expect(Buffer.compare(Buffer.from(wb.toBytes()), Buffer.from(beforeBytes)), entry.name).toBe(
				0,
			)
		}
	})

	test('saved-source package state journals are lossy when inverse ops cannot restore bytes', async () => {
		const cases: readonly {
			readonly name: string
			readonly seedOps: readonly Operation[]
			readonly ops: readonly Operation[]
			readonly cleanCalcState?: boolean
			readonly opName: Operation['op']
			readonly refs: readonly string[]
		}[] = [
			{
				name: 'addSheet',
				seedOps: [],
				ops: [{ op: 'addSheet', name: 'Added' }],
				cleanCalcState: true,
				opName: 'addSheet',
				refs: ['sheet:Added'],
			},
			{
				name: 'renameSheet',
				seedOps: [],
				ops: [{ op: 'renameSheet', sheet: 'Sheet1', newName: 'Renamed' }],
				cleanCalcState: true,
				opName: 'renameSheet',
				refs: ['sheet:Sheet1', 'sheet:Renamed'],
			},
			{
				name: 'moveSheet',
				seedOps: [{ op: 'addSheet', name: 'Second' }],
				ops: [{ op: 'moveSheet', sheet: 'Second', position: 0 }],
				cleanCalcState: true,
				opName: 'moveSheet',
				refs: ['sheet:Second'],
			},
			{
				name: 'copySheet',
				seedOps: [],
				ops: [{ op: 'copySheet', sheet: 'Sheet1', newName: 'Copy' }],
				cleanCalcState: true,
				opName: 'copySheet',
				refs: ['sheet:Sheet1', 'sheet:Copy'],
			},
			{
				name: 'deleteSheet',
				seedOps: [{ op: 'addSheet', name: 'Empty' }],
				ops: [{ op: 'deleteSheet', sheet: 'Empty' }],
				cleanCalcState: true,
				opName: 'deleteSheet',
				refs: ['sheet:Empty'],
			},
			{
				name: 'setDocumentProperties',
				seedOps: [
					{
						op: 'setDocumentProperties',
						mode: 'replace',
						properties: {
							core: { title: 'Before' },
							app: { company: 'Ascend' },
							custom: [{ name: 'Reviewed', value: false, type: 'bool' }],
						},
					},
				],
				ops: [
					{
						op: 'setDocumentProperties',
						mode: 'replace',
						properties: {
							core: { title: 'After' },
							app: { company: 'Changed' },
							custom: [{ name: 'Reviewed', value: true, type: 'bool' }],
						},
					},
				],
				opName: 'setDocumentProperties',
				refs: ['workbook:documentProperties'],
			},
			{
				name: 'setWorkbookProperties',
				seedOps: [{ op: 'setWorkbookProperties', properties: { date1904: false } }],
				ops: [{ op: 'setWorkbookProperties', properties: { date1904: true } }],
				cleanCalcState: true,
				opName: 'setWorkbookProperties',
				refs: ['workbook:properties'],
			},
			{
				name: 'setCells',
				seedOps: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
				ops: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] }],
				cleanCalcState: true,
				opName: 'setCells',
				refs: ['Sheet1!A1'],
			},
			{
				name: 'setFormula',
				seedOps: [
					{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
					{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
				],
				ops: [{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*3' }],
				cleanCalcState: true,
				opName: 'setFormula',
				refs: ['Sheet1!B1'],
			},
			{
				name: 'fillFormula',
				seedOps: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [
							{ ref: 'A1', value: 1 },
							{ ref: 'A2', value: 2 },
						],
					},
					{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
				],
				ops: [{ op: 'fillFormula', sheet: 'Sheet1', range: 'B1:B2', formula: 'A1*2' }],
				cleanCalcState: true,
				opName: 'fillFormula',
				refs: ['Sheet1!B1:B2'],
			},
			{
				name: 'clearRange',
				seedOps: [{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }],
				ops: [{ op: 'clearRange', sheet: 'Sheet1', range: 'A1', what: 'all' }],
				cleanCalcState: true,
				opName: 'clearRange',
				refs: ['Sheet1!A1'],
			},
			{
				name: 'setRichText',
				seedOps: [{ op: 'setRichText', sheet: 'Sheet1', ref: 'A1', runs: [{ text: 'one' }] }],
				ops: [{ op: 'setRichText', sheet: 'Sheet1', ref: 'A1', runs: [{ text: 'two' }] }],
				cleanCalcState: true,
				opName: 'setRichText',
				refs: ['Sheet1!A1'],
			},
			{
				name: 'copyRange',
				seedOps: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [
							{ ref: 'A1', value: 1 },
							{ ref: 'D1', value: 9 },
						],
					},
				],
				ops: [{ op: 'copyRange', sheet: 'Sheet1', source: 'A1', target: 'D1', mode: 'all' }],
				cleanCalcState: true,
				opName: 'copyRange',
				refs: ['Sheet1!A1', 'Sheet1!D1'],
			},
			{
				name: 'moveRange',
				seedOps: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [
							{ ref: 'A1', value: 1 },
							{ ref: 'D1', value: 9 },
						],
					},
				],
				ops: [{ op: 'moveRange', sheet: 'Sheet1', source: 'A1', target: 'D1', mode: 'all' }],
				cleanCalcState: true,
				opName: 'moveRange',
				refs: ['Sheet1!A1', 'Sheet1!D1'],
			},
			{
				name: 'sortRange',
				seedOps: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [
							{ ref: 'A1', value: 'Key' },
							{ ref: 'B1', value: 'Qty' },
							{ ref: 'A2', value: 'b' },
							{ ref: 'B2', value: 2 },
							{ ref: 'A3', value: 'a' },
							{ ref: 'B3', value: 1 },
						],
					},
				],
				ops: [{ op: 'sortRange', sheet: 'Sheet1', range: 'A1:B3', by: [{ column: 'Qty' }] }],
				cleanCalcState: true,
				opName: 'sortRange',
				refs: ['Sheet1!A1:B3'],
			},
			{
				name: 'renameTable',
				seedOps: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [
							{ ref: 'A1', value: 'Qty' },
							{ ref: 'A2', value: 1 },
						],
					},
					{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:A2', name: 'Sales', hasHeaders: true },
				],
				ops: [{ op: 'renameTable', table: 'Sales', newName: 'Revenue' }],
				cleanCalcState: true,
				opName: 'renameTable',
				refs: ['table:Sales', 'table:Revenue'],
			},
			{
				name: 'setTableColumn',
				seedOps: [
					{
						op: 'setCells',
						sheet: 'Sheet1',
						updates: [
							{ ref: 'A1', value: 'Qty' },
							{ ref: 'A2', value: 1 },
						],
					},
					{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:A2', name: 'Sales', hasHeaders: true },
				],
				ops: [{ op: 'setTableColumn', table: 'Sales', column: 'Qty', newName: 'Units' }],
				cleanCalcState: true,
				opName: 'setTableColumn',
				refs: ['table:Sales'],
			},
			{
				name: 'setDefinedName',
				seedOps: [{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$A$1' }],
				ops: [{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$B$1' }],
				cleanCalcState: true,
				opName: 'setDefinedName',
				refs: ['name:Budget'],
			},
			{
				name: 'deleteDefinedName',
				seedOps: [{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$A$1' }],
				ops: [{ op: 'deleteDefinedName', name: 'Budget' }],
				cleanCalcState: true,
				opName: 'deleteDefinedName',
				refs: ['name:Budget'],
			},
		]

		for (const entry of cases) {
			const seeded = AscendWorkbook.create()
			applyExact(seeded, entry.seedOps)
			if (entry.cleanCalcState) expect(seeded.recalc().errors, entry.name).toEqual([])
			const beforeBytes = seeded.toBytes()
			const wb = await AscendWorkbook.open(beforeBytes)
			const changed = wb.apply(entry.ops, { journal: true })
			expect(changed.errors, entry.name).toEqual([])
			expect(changed.journal?.supported, entry.name).toBe(true)
			expect(changed.journal?.exact, entry.name).toBe(false)
			expect(changed.journal?.issues, entry.name).toEqual([
				{
					code: 'LOSSY_INVERSE',
					message: `${entry.opName} changes saved package state that public inverse operations cannot restore byte-for-byte`,
					surface: 'package-parts',
					reason: 'package-part-preservation',
					refs: entry.refs,
				},
			])

			const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
			expect(undo.errors, entry.name).toEqual([])
			expect(
				Buffer.compare(Buffer.from(wb.toBytes()), Buffer.from(beforeBytes)),
				entry.name,
			).not.toBe(0)
		}
	})
})

function lossyDataValidationDefaultJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.dataValidations.push({ sqref: 'A1:A1', type: 'whole', formula1: '1' })
	return applyJournal(wb, [
		{
			op: 'setDataValidation',
			sheet: 'Sheet1',
			range: 'A1:A1',
			rule: { type: 'whole', formula1: '2' },
		},
	])
}

function lossyDataValidationDeleteOrderJournal(): MutationJournal {
	const wb = dataValidationOrderWorkbook()
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'deleteDataValidation', sheet: 'Sheet1', range: 'A1:A1' },
	])
}

function dataValidationOrderWorkbook(): AscendWorkbook {
	const wb = AscendWorkbook.create()
	applyExact(wb, [
		{
			op: 'setDataValidation',
			sheet: 'Sheet1',
			range: 'A1:A1',
			rule: { type: 'whole', formula1: '1', allowBlank: true, showErrorMessage: true },
		},
		{
			op: 'setDataValidation',
			sheet: 'Sheet1',
			range: 'B1:B1',
			rule: { type: 'whole', formula1: '2', allowBlank: true, showErrorMessage: true },
		},
	])
	return wb
}

function lossyDataValidationDuplicateDeleteJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.dataValidations.push(
		{ sqref: 'A1:A1', type: 'whole', formula1: '1', allowBlank: true, showErrorMessage: true },
		{ sqref: 'A1:A1', type: 'whole', formula1: '2', allowBlank: true, showErrorMessage: true },
	)
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'deleteDataValidation', sheet: 'Sheet1', range: 'A1:A1' },
	])
}

function lossyDataValidationDuplicateMoveJournal(): MutationJournal {
	const wb = dataValidationDuplicateWorkbook()
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'moveRange', sheet: 'Sheet1', source: 'A1:A1', target: 'D1', mode: 'validations' },
	])
}

function dataValidationDuplicateWorkbook(): AscendWorkbook {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.dataValidations.push(
		{ sqref: 'A1:A1', type: 'whole', formula1: '1', allowBlank: true, showErrorMessage: true },
		{ sqref: 'A1:A1', type: 'whole', formula1: '2', allowBlank: true, showErrorMessage: true },
	)
	return wb
}

function lossyLegacyCommentDrawingJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.comments.set('A1', {
		text: 'legacy note',
		legacyDrawing: { shapeId: '_x0000_s1025' },
	})
	return applyJournal(wb, [{ op: 'deleteComment', sheet: 'Sheet1', ref: 'A1' }])
}

function lossySetCommentLegacyDrawingJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.comments.set('A1', {
		text: 'legacy note',
		legacyDrawing: { shapeId: '_x0000_s1025' },
	})
	return applyJournal(wb, [{ op: 'setComment', sheet: 'Sheet1', ref: 'A1', text: 'updated' }])
}

function lossyCreatedLayoutJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	return applyJournal(wb, [
		{ op: 'setRowHeight', sheet: 'Sheet1', row: 2, height: 32 },
		{ op: 'setColWidth', sheet: 'Sheet1', col: 1, width: 20 },
	])
}

function lossyThreadedCommentSelectorJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.threadedComments.push(
		{ ref: 'B2', text: 'first', id: 'tc-1' },
		{ ref: 'B2', text: 'second', id: 'tc-2' },
	)
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'setThreadedComment', sheet: 'Sheet1', ref: 'B2', text: 'ambiguous' },
	])
}

function lossyDrawingSelectorJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.drawingObjectRefs.push(
		{
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'textBox',
			id: 1,
			name: 'Duplicate',
			text: 'First',
		},
		{
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'shape',
			id: 2,
			name: 'Duplicate',
		},
	)
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'setDrawingText', sheet: 'Sheet1', name: 'Duplicate', text: 'Updated' },
	])
}

function lossyChartSeriesUnsetJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	wb.getWorkbookModel().chartParts.push({
		partPath: 'xl/charts/chart1.xml',
		sheetName: 'Sheet1',
		series: [{ nameText: 'Actual', valueRef: 'Sheet1!$B$2:$B$4' }],
	})
	return buildMutationJournal(wb.getWorkbookModel(), [
		{
			op: 'setChartSeriesSource',
			partPath: 'xl/charts/chart1.xml',
			seriesIndex: 0,
			nameRef: 'Sheet1!$B$1',
			valueRef: 'Sheet1!$C$2:$C$4',
		},
	])
}

function lossyPivotCacheUnsetJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	wb.getWorkbookModel().pivotCaches.push({
		partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
		cacheId: 34,
		sourceRef: 'A1:D10',
		fields: [],
	})
	return buildMutationJournal(wb.getWorkbookModel(), [
		{
			op: 'setPivotCache',
			cacheId: 34,
			sourceSheet: 'RawData',
			sourceRef: 'A1:E20',
			refreshOnLoad: true,
		},
	])
}

function lossyConditionalFormatOrderJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	applyExact(wb, [
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A1',
			rule: { type: 'expression', formula: 'A1>0', priority: 1 },
		},
	])
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'deleteConditionalFormat', sheet: 'Sheet1' },
	])
}

function lossyConditionalFormatReplacementOrderJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	applyExact(wb, [
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A2',
			rule: { type: 'expression', formula: 'A1>0', priority: 1 },
		},
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'B1:B2',
			rule: { type: 'expression', formula: 'B1>0', priority: 2 },
		},
	])
	return buildMutationJournal(wb.getWorkbookModel(), [
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A2',
			rule: { type: 'expression', formula: 'A1>5', priority: 1 },
		},
	])
}

function lossyConditionalFormatDeleteOrderJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	applyExact(wb, [
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'A1:A2',
			rule: { type: 'expression', formula: 'A1>0', priority: 1 },
		},
		{
			op: 'setConditionalFormat',
			sheet: 'Sheet1',
			range: 'B1:B2',
			rule: { type: 'expression', formula: 'B1>0', priority: 2 },
		},
	])
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'A1:A2' },
	])
}

function lossyConditionalFormatDuplicateDeleteJournal(): MutationJournal {
	const wb = conditionalFormatDuplicateWorkbook()
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'A1:A1' },
	])
}

function lossyConditionalFormatDuplicateMoveJournal(): MutationJournal {
	const wb = conditionalFormatDuplicateWorkbook()
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'moveRange', sheet: 'Sheet1', source: 'A1:A1', target: 'D1', mode: 'formats' },
	])
}

function conditionalFormatDuplicateWorkbook(): AscendWorkbook {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.conditionalFormats.push(
		{ sqref: 'A1:A1', rules: [{ type: 'expression', formulas: ['A1>0'] }] },
		{ sqref: 'A1:A1', rules: [{ type: 'expression', formulas: ['A1<0'] }] },
	)
	return wb
}

function lossyMergeDuplicateMoveJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	const duplicateMerge = { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }
	sheet.merges.push(duplicateMerge, duplicateMerge)
	return buildMutationJournal(wb.getWorkbookModel(), [
		{ op: 'moveRange', sheet: 'Sheet1', source: 'A1:B1', target: 'D1', mode: 'formats' },
	])
}

function lossyAutoFilterJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.autoFilter = {
		ref: 'A1:B10',
		uid: '{filter-uid}',
		columns: [{ colId: 0, kind: 'filters', values: ['Open'], hiddenButton: true }],
		sortState: {
			ref: 'A1:B10',
			caseSensitive: true,
			conditions: [{ ref: 'A2:A10', descending: true }, { ref: 'B2:B10' }],
		},
	}
	return applyJournal(wb, [{ op: 'clearAutoFilter', sheet: 'Sheet1' }])
}

function lossyPageSetupJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	return applyJournal(wb, [
		{ op: 'setPageSetup', sheet: 'Sheet1', setup: { orientation: 'landscape' } },
	])
}

function lossyX14TransferJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	const sheet = wb.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.x14DataValidations.push({
		index: 0,
		sqref: 'A1:A1',
		type: 'list',
		formula1: '"A,B"',
	})
	return applyJournal(wb, [
		{ op: 'copyRange', sheet: 'Sheet1', source: 'A1', target: 'D1', mode: 'validations' },
	])
}

function lossyOverlappingMoveRangeJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	applyExact(wb, [
		{
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [
				{ ref: 'A1', value: 1 },
				{ ref: 'A2', value: 2 },
			],
		},
	])
	return applyJournal(wb, [
		{ op: 'moveRange', sheet: 'Sheet1', source: 'A1:A2', target: 'A2', mode: 'all' },
	])
}

function lossyStyleRegistryJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	return applyJournal(wb, [
		{
			op: 'setStyle',
			sheet: 'Sheet1',
			range: 'A1:A1',
			style: { font: { bold: true } },
		},
	])
}

function lossyDeletedSheetJournal(): MutationJournal {
	const wb = AscendWorkbook.create()
	applyExact(wb, [{ op: 'addSheet', name: 'Data' }])
	const sheet = wb.getWorkbookModel().getSheet('Data')
	if (!sheet) throw new Error('missing sheet')
	sheet.cells.set(0, 0, {
		value: numberValue(1),
		formula: null,
		styleId: DEFAULT_STYLE_ID,
	})
	return applyJournal(wb, [{ op: 'deleteSheet', sheet: 'Data' }])
}

function seedSharedFormulaPair(workbook: AscendWorkbook, row: number, col: number) {
	const sheet = workbook.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	const column = String.fromCharCode(65 + col)
	const nextColumn = String.fromCharCode(66 + col)
	const topRef = `${column}${row + 1}`
	const bottomRef = `${column}${row + 2}`
	sheet.cells.set(row, col + 1, {
		value: numberValue(10),
		formula: null,
		styleId: DEFAULT_STYLE_ID,
	})
	sheet.cells.set(row + 1, col + 1, {
		value: numberValue(20),
		formula: null,
		styleId: DEFAULT_STYLE_ID,
	})
	sheet.cells.set(row, col, {
		value: numberValue(20),
		formula: `${nextColumn}${row + 1}*2`,
		styleId: DEFAULT_STYLE_ID,
		formulaInfo: {
			kind: 'shared',
			sharedIndex: `journal-range-${topRef}`,
			isMaster: true,
			masterRef: topRef,
			ref: `${topRef}:${bottomRef}`,
		},
	})
	sheet.cells.set(row + 1, col, {
		value: numberValue(40),
		formula: null,
		styleId: DEFAULT_STYLE_ID,
		formulaInfo: {
			kind: 'shared',
			sharedIndex: `journal-range-${topRef}`,
			isMaster: false,
			masterRef: topRef,
		},
	})
	return sheet
}

function seedDynamicArrayFootprint(workbook: AscendWorkbook, col: number) {
	const sheet = workbook.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	const column = String.fromCharCode(65 + col)
	const anchorRef = `Sheet1!${column}1`
	const ref = `${column}1:${column}3`
	sheet.cells.set(0, col, {
		value: numberValue(1),
		formula: 'SEQUENCE(3)',
		styleId: DEFAULT_STYLE_ID,
		formulaInfo: { kind: 'dynamicArray', metadataIndex: 1, collapsed: false },
	})
	sheet.cells.set(1, col, {
		value: numberValue(2),
		formula: null,
		styleId: DEFAULT_STYLE_ID,
		formulaInfo: { kind: 'spill', anchorRef, ref, isAnchor: false },
	})
	sheet.cells.set(2, col, {
		value: numberValue(3),
		formula: null,
		styleId: DEFAULT_STYLE_ID,
		formulaInfo: { kind: 'spill', anchorRef, ref, isAnchor: false },
	})
	return sheet
}

function seedDataTableFormula(workbook: AscendWorkbook) {
	const sheet = workbook.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	for (let row = 0; row < 5; row++) {
		sheet.cells.set(row, 0, {
			value: numberValue(row + 1),
			formula: null,
			styleId: DEFAULT_STYLE_ID,
		})
	}
	sheet.cells.set(2, 2, {
		value: numberValue(10),
		formula: null,
		styleId: DEFAULT_STYLE_ID,
		formulaInfo: { kind: 'dataTable', ref: 'C3:C5', dt2D: false, dtr: true, r1: 'A1' },
	})
	sheet.cells.set(3, 2, { value: numberValue(20), formula: null, styleId: DEFAULT_STYLE_ID })
	sheet.cells.set(4, 2, { value: numberValue(30), formula: null, styleId: DEFAULT_STYLE_ID })
	return sheet
}

function seedBlockedSpill(workbook: AscendWorkbook) {
	const sheet = workbook.getWorkbookModel().getSheet('Sheet1')
	if (!sheet) throw new Error('missing sheet')
	sheet.cells.set(0, 0, {
		value: errorValue('#SPILL!'),
		formula: 'SEQUENCE(3)',
		styleId: DEFAULT_STYLE_ID,
		formulaInfo: {
			kind: 'blockedSpill',
			anchorRef: 'Sheet1!A1',
			ref: 'A1:A3',
			blockingRefs: ['A2'],
		},
	})
	sheet.cells.set(1, 0, {
		value: stringValue('blocker'),
		formula: null,
		styleId: DEFAULT_STYLE_ID,
	})
	return sheet
}

function journalEvidence(wb: AscendWorkbook): object {
	const snapshot = wb.snapshot()
	const { timestamp: _timestamp, ...stableSnapshot } = snapshot
	const model = wb.getWorkbookModel()
	return {
		snapshot: stableSnapshot,
		styleRegistry: Array.from({ length: model.styles.size }, (_, index) => model.styles.get(index)),
		workbookProperties: model.workbookProperties,
		documentProperties: wb.inspect().documentProperties,
		workbookViews: wb.inspect().workbookViews,
		calcSettings: model.calcSettings,
		themeSummary: wb.inspect().themeSummary,
		sheetMetadata: model.sheets.map((sheet) => ({
			name: sheet.name,
			merges: sheet.merges,
			dataValidations: sheet.dataValidations,
			conditionalFormats: sheet.conditionalFormats,
			x14DataValidations: sheet.x14DataValidations,
			x14ConditionalFormats: sheet.x14ConditionalFormats,
			comments: [...sheet.comments.entries()],
			threadedComments: sheet.threadedComments,
			hyperlinks: [...sheet.hyperlinks.entries()],
			tables: sheet.tables,
			tabColor: sheet.tabColor,
			protection: sheet.protection,
			state: sheet.state,
			outlinePr: sheet.outlinePr,
			sheetFormatPr: sheet.sheetFormatPr,
			rowHeights: [...sheet.rowHeights.entries()],
			colWidths: [...sheet.colWidths.entries()],
			rowDefs: [...sheet.rowDefs.entries()],
			colDefs: sheet.colDefs,
			pageSetup: sheet.pageSetup,
			pageMargins: sheet.pageMargins,
			printOptions: sheet.printOptions,
			printArea: sheet.printArea,
			frozenRows: sheet.frozenRows,
			frozenCols: sheet.frozenCols,
			formulaBindings: [...sheet.cells.iterate()]
				.filter(([, , cell]) => cell.formulaInfo)
				.map(([row, col, cell]) => ({
					ref: `${indexToColumn(col)}${row + 1}`,
					formulaInfo: cell.formulaInfo,
				})),
		})),
		chartParts: model.chartParts,
		chartSheets: model.chartSheets,
		pivotCaches: model.pivotCaches,
		pivotTables: model.pivotTables,
	}
}

function applyExact(workbook: AscendWorkbook, ops: readonly Operation[]): void {
	const result = workbook.apply(ops)
	expect(result.errors).toEqual([])
}

function applyJournal(workbook: AscendWorkbook, ops: readonly Operation[]): MutationJournal {
	const result = workbook.apply(ops, { journal: true })
	expect(result.errors).toEqual([])
	const journal = result.journal
	expect(journal).toBeDefined()
	if (!journal) throw new Error('missing journal')
	return journal
}

function collectFormulaInfoRefs(workbook: AscendWorkbook, sheetName: string): string[] {
	const sheet = workbook.getWorkbookModel().getSheet(sheetName)
	if (!sheet) throw new Error(`missing sheet ${sheetName}`)
	const refs: string[] = []
	for (const [row, col, cell] of sheet.cells.iterate()) {
		if (cell.formulaInfo) refs.push(`${sheetName}!${indexToColumn(col)}${row + 1}`)
	}
	return refs
}

function journalIssueRefs(
	journal: MutationJournal,
	surface: MutationJournalSurface,
): readonly string[] {
	return journal.issues
		.filter((issue) => issue.surface === surface)
		.flatMap((issue) => issue.refs ?? [])
		.sort()
}

function formulaInfoRefsIn(
	workbook: AscendWorkbook,
	sheetName: string,
	refs: readonly string[],
): string[] {
	const sheet = workbook.getWorkbookModel().getSheet(sheetName)
	if (!sheet) throw new Error(`missing sheet ${sheetName}`)
	return refs
		.filter((ref) => {
			const parsed = parseA1(ref)
			return sheet.cells.get(parsed.row, parsed.col)?.formulaInfo
		})
		.map((ref) => `${sheetName}!${ref}`)
}

function cellFormulas(
	workbook: AscendWorkbook,
	sheetName: string,
	refs: readonly string[],
): Record<string, string | null | undefined> {
	const sheet = workbook.getWorkbookModel().getSheet(sheetName)
	if (!sheet) throw new Error(`missing sheet ${sheetName}`)
	return Object.fromEntries(
		refs.map((ref) => {
			const parsed = parseA1(ref)
			return [ref, sheet.cells.get(parsed.row, parsed.col)?.formula]
		}),
	)
}
