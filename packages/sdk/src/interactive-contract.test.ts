import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'
import { buildMutationJournal } from './journal.ts'

describe('interactive client contract', () => {
	test('apply returns dirty regions and monotonic generation tokens', () => {
		const wb = AscendWorkbook.create()

		const valueEdit = wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'B2', value: 10 },
					{ ref: 'D4', value: 20 },
				],
			},
		])
		expect(valueEdit.errors).toEqual([])
		expect(valueEdit.dirtyRegions).toEqual([
			{
				sheet: 'Sheet1',
				range: 'B2:D4',
				refs: ['Sheet1!B2', 'Sheet1!D4'],
			},
		])
		expect(valueEdit.generations).toEqual({
			workbook: 1,
			sheetMetadata: 0,
			formulas: 1,
			styles: 0,
		})

		const styleEdit = wb.apply([
			{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'B2:B2', format: '0.00' },
		])
		expect(styleEdit.errors).toEqual([])
		expect(styleEdit.dirtyRegions).toEqual([
			{ sheet: 'Sheet1', range: 'B2:B2', refs: ['Sheet1!B2'] },
		])
		expect(styleEdit.generations).toEqual({
			workbook: 2,
			sheetMetadata: 0,
			formulas: 1,
			styles: 1,
		})

		const structuralEdit = wb.apply([{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 }])
		expect(structuralEdit.errors).toEqual([])
		expect(structuralEdit.generations).toEqual({
			workbook: 3,
			sheetMetadata: 1,
			formulas: 2,
			styles: 1,
		})
	})

	test('recalc returns changed dirty regions and advances formula generation once', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [{ ref: 'A1', value: 1 }],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1*2' },
		])
		const before = wb.recalc()
		expect(before.errors).toEqual([])

		const apply = wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 5 }] },
		])
		const recalc = wb.recalc()

		expect(apply.dirtyRegions).toEqual([{ sheet: 'Sheet1', range: 'A1:A1', refs: ['Sheet1!A1'] }])
		expect(recalc.changed).toEqual(['Sheet1!B1'])
		expect(recalc.dirtyRegions).toEqual([{ sheet: 'Sheet1', range: 'B1:B1', refs: ['Sheet1!B1'] }])
		expect(recalc.generations.formulas).toBeGreaterThan(apply.generations.formulas)
		expect(recalc.generations.workbook).toBeGreaterThan(apply.generations.workbook)
	})

	test('apply can return a reversible journal for supported interactive edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }])

		const apply = wb.apply(
			[
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
				{ op: 'setComment', sheet: 'Sheet1', ref: 'B2', text: 'review', author: 'agent' },
			],
			{ journal: true },
		)

		expect(apply.errors).toEqual([])
		expect(apply.journal?.supported).toBe(true)
		expect(apply.journal?.exact).toBe(true)
		expect(apply.journal?.inverseOps).toEqual([
			{ op: 'deleteComment', sheet: 'Sheet1', ref: 'B2' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
		])

		const undo = wb.apply(apply.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A1')?.value).toEqual({ kind: 'number', value: 1 })
		expect(wb.sheet('Sheet1')?.comment('B2')).toBeUndefined()
	})

	test('preview journals expose lossy formula preimages without mutating the workbook', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '1+1' }])

		const preview = wb.preview(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] }],
			{ journal: true },
		)

		expect(preview.wouldSucceed).toBe(true)
		expect(preview.journal?.supported).toBe(true)
		expect(preview.journal?.exact).toBe(false)
		expect(preview.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Formula cache for Sheet1!A1 cannot be restored with public operations',
				refs: ['Sheet1!A1'],
			},
		])
		expect(preview.journal?.inverseOps).toEqual([
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '1+1' },
		])
		expect(wb.sheet('Sheet1')?.cell('A1')?.formula).toBe('1+1')
	})

	test('journal inverse ops preserve internal order when restoring cleared cell state', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'C3', value: 12 }] }])
		wb.apply([{ op: 'setNumberFormat', sheet: 'Sheet1', range: 'C3:C3', format: '0.00' }])

		const cleared = wb.apply([{ op: 'clearRange', sheet: 'Sheet1', range: 'C3:C3', what: 'all' }], {
			journal: true,
		})
		expect(cleared.journal?.exact).toBe(true)

		const undo = wb.apply(cleared.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('C3')?.value).toEqual({ kind: 'number', value: 12 })
		expect(wb.cellStyle('Sheet1!C3')?.numberFormat).toBe('0.00')
	})

	test('journal inverse ops restore merge and unmerge metadata', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'mergeCells', sheet: 'Sheet1', range: 'A1:B2' }])

		const unmerged = wb.apply([{ op: 'unmergeCells', sheet: 'Sheet1', range: 'A1:B2' }], {
			journal: true,
		})

		expect(unmerged.journal?.exact).toBe(true)
		expect(unmerged.journal?.inverseOps).toEqual([
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'A1:B2' },
		])

		const undoUnmerge = wb.apply(unmerged.journal?.inverseOps ?? [], { transaction: true })
		expect(undoUnmerge.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.merges).toEqual([
			{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
		])

		const wb2 = AscendWorkbook.create()
		const merged = wb2.apply([{ op: 'mergeCells', sheet: 'Sheet1', range: 'C1:D1' }], {
			journal: true,
		})
		expect(merged.journal?.inverseOps).toEqual([
			{ op: 'unmergeCells', sheet: 'Sheet1', range: 'C1:D1' },
		])
	})

	test('journal inverse ops restore data validation changes', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A10',
				rule: {
					type: 'list',
					formula1: '"Open,Closed"',
					allowBlank: true,
					showErrorMessage: true,
				},
			},
		])

		const changed = wb.apply(
			[
				{
					op: 'setDataValidation',
					sheet: 'Sheet1',
					range: 'A1:A10',
					rule: { type: 'whole', operator: 'greaterThan', formula1: '0' },
				},
			],
			{ journal: true },
		)

		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A1:A10',
				rule: {
					type: 'list',
					formula1: '"Open,Closed"',
					allowBlank: true,
					showErrorMessage: true,
				},
			},
		])

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.getDataValidations()).toEqual([
			{ type: 'list', formula: '"Open,Closed"', range: 'A1:A10' },
		])
	})

	test('journal inverse ops restore simple worksheet auto filters', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C10' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C10', column: 1, values: ['Open'] },
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C10',
				sortRef: 'A2:C10',
				sortBy: 'B2:B10',
				descending: true,
			},
		])

		const cleared = wb.apply([{ op: 'clearAutoFilter', sheet: 'Sheet1' }], { journal: true })

		expect(cleared.journal?.exact).toBe(true)
		expect(cleared.journal?.inverseOps).toEqual([
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C10' },
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C10',
				column: 1,
				values: ['Open'],
			},
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:C10',
				sortRef: 'A2:C10',
				sortBy: 'B2:B10',
				descending: true,
			},
		])

		const undo = wb.apply(cleared.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.autoFilter).toEqual({
			ref: 'A1:C10',
			columns: [{ colId: 1, kind: 'filters', values: ['Open'] }],
			sortState: {
				ref: 'A2:C10',
				conditions: [{ ref: 'B2:B10', descending: true }],
			},
		})

		const wb2 = AscendWorkbook.create()
		const created = wb2.apply([{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'D1:E5' }], {
			journal: true,
		})
		expect(created.journal?.inverseOps).toEqual([{ op: 'clearAutoFilter', sheet: 'Sheet1' }])
	})

	test('journal inverse ops restore deleted hyperlink metadata exactly', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'b2',
				url: 'https://example.com/report',
				display: 'Report',
				tooltip: 'Open report',
			},
		])

		const deleted = wb.apply([{ op: 'deleteHyperlink', sheet: 'Sheet1', ref: 'B2' }], {
			journal: true,
		})

		expect(deleted.errors).toEqual([])
		expect(deleted.journal?.exact).toBe(true)
		expect(deleted.journal?.inverseOps).toEqual([
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'B2',
				url: 'https://example.com/report',
				display: 'Report',
				tooltip: 'Open report',
			},
		])
		expect(wb.sheet('Sheet1')?.getHyperlinks()).toEqual([])

		const undo = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.getHyperlinks()).toEqual([
			{
				ref: 'B2',
				target: 'https://example.com/report',
				display: 'Report',
				tooltip: 'Open report',
			},
		])
	})

	test('journal inverse ops restore threaded comment text without losing metadata', () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.threadedComments.push(
			{
				ref: 'B2',
				text: 'original',
				id: 'tc-1',
				personId: 'person-1',
				author: 'Analyst',
				dateTime: '2026-05-13T10:00:00Z',
				done: false,
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
			{
				ref: 'B2',
				text: 'follow-up',
				id: 'tc-2',
				parentId: 'tc-1',
				personId: 'person-2',
				author: 'Reviewer',
				dateTime: '2026-05-13T10:05:00Z',
				done: true,
				partPath: 'xl/threadedComments/threadedComment1.xml',
			},
		)

		const changed = wb.apply(
			[
				{
					op: 'setThreadedComment',
					sheet: 'Sheet1',
					partPath: 'xl/threadedComments/threadedComment1.xml',
					threadedCommentId: 'tc-2',
					text: 'resolved by agent',
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setThreadedComment',
				sheet: 'Sheet1',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				threadedCommentId: 'tc-2',
				text: 'follow-up',
			},
		])
		const changedThread = wb.getWorkbookModel().sheets[0]?.threadedComments[1]
		expect(changedThread).toMatchObject({
			text: 'resolved by agent',
			id: 'tc-2',
			parentId: 'tc-1',
			personId: 'person-2',
			author: 'Reviewer',
			done: true,
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.getWorkbookModel().sheets[0]?.threadedComments[1]).toEqual({
			ref: 'B2',
			text: 'follow-up',
			id: 'tc-2',
			parentId: 'tc-1',
			personId: 'person-2',
			author: 'Reviewer',
			dateTime: '2026-05-13T10:05:00Z',
			done: true,
			partPath: 'xl/threadedComments/threadedComment1.xml',
		})
	})

	test('threaded comment journals do not claim exact rollback for ambiguous selectors', () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.threadedComments.push(
			{ ref: 'B2', text: 'first', id: 'tc-1' },
			{ ref: 'B2', text: 'second', id: 'tc-2' },
		)

		const journal = buildMutationJournal(wb.getWorkbookModel(), [
			{ op: 'setThreadedComment', sheet: 'Sheet1', ref: 'B2', text: 'ambiguous' },
		])

		expect(journal.supported).toBe(true)
		expect(journal.exact).toBe(false)
		expect(journal.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Threaded comment selector on Sheet1 cannot be resolved exactly',
			},
		])
		expect(journal.inverseOps).toEqual([])
		expect(sheet?.threadedComments.map((comment) => comment.text)).toEqual(['first', 'second'])
	})

	test('journal inverse ops restore drawing text without losing relationship metadata', () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.drawingObjectRefs.push({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'graphicFrame',
			id: 7,
			name: 'Chart Callout',
			description: 'Revenue callout',
			text: 'Original',
			anchor: {
				kind: 'absolute',
				x: 1000,
				y: 2000,
				cx: 3000,
				cy: 4000,
			},
			relIds: ['rIdChart1'],
			relationshipRefs: [
				{
					id: 'rIdChart1',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
					target: '../charts/chart1.xml',
				},
			],
		})

		const changed = wb.apply(
			[
				{
					op: 'setDrawingText',
					sheet: 'Sheet1',
					drawingPartPath: 'xl/drawings/drawing1.xml',
					id: 7,
					text: 'Updated',
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setDrawingText',
				sheet: 'Sheet1',
				drawingPartPath: 'xl/drawings/drawing1.xml',
				id: 7,
				text: 'Original',
			},
		])
		expect(wb.getWorkbookModel().sheets[0]?.drawingObjectRefs[0]).toMatchObject({
			id: 7,
			name: 'Chart Callout',
			text: 'Updated',
			relIds: ['rIdChart1'],
		})

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.getWorkbookModel().sheets[0]?.drawingObjectRefs[0]).toEqual({
			drawingPartPath: 'xl/drawings/drawing1.xml',
			kind: 'graphicFrame',
			id: 7,
			name: 'Chart Callout',
			description: 'Revenue callout',
			text: 'Original',
			anchor: {
				kind: 'absolute',
				x: 1000,
				y: 2000,
				cx: 3000,
				cy: 4000,
			},
			relIds: ['rIdChart1'],
			relationshipRefs: [
				{
					id: 'rIdChart1',
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
					target: '../charts/chart1.xml',
				},
			],
		})
	})

	test('drawing text journals do not claim exact rollback for ambiguous or non-text selectors', () => {
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.drawingObjectRefs.push(
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

		const ambiguous = buildMutationJournal(wb.getWorkbookModel(), [
			{ op: 'setDrawingText', sheet: 'Sheet1', name: 'Duplicate', text: 'Updated' },
		])
		expect(ambiguous.supported).toBe(true)
		expect(ambiguous.exact).toBe(false)
		expect(ambiguous.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Drawing object selector on Sheet1 cannot be resolved to editable text exactly',
			},
		])
		expect(ambiguous.inverseOps).toEqual([])

		const noText = buildMutationJournal(wb.getWorkbookModel(), [
			{ op: 'setDrawingText', sheet: 'Sheet1', id: 2, text: 'Updated' },
		])
		expect(noText.supported).toBe(true)
		expect(noText.exact).toBe(false)
		expect(noText.inverseOps).toEqual([])
		expect(sheet?.drawingObjectRefs.map((object) => object.text)).toEqual(['First', undefined])
	})

	test('journal inverse ops restore chart series source refs without losing chart metadata', () => {
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'lineChart',
			title: 'Revenue',
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$B$2:$B$4',
				},
				{
					nameText: 'Plan',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$C$2:$C$4',
				},
			],
		})

		const changed = wb.apply(
			[
				{
					op: 'setChartSeriesSource',
					partPath: 'xl/charts/chart1.xml',
					seriesIndex: 0,
					nameRef: 'Sheet1!$D$1',
					categoryRef: 'Sheet1!$A$2:$A$10',
					valueRef: 'Sheet1!$D$2:$D$10',
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setChartSeriesSource',
				partPath: 'xl/charts/chart1.xml',
				seriesIndex: 0,
				nameRef: 'Sheet1!$B$1',
				categoryRef: 'Sheet1!$A$2:$A$4',
				valueRef: 'Sheet1!$B$2:$B$4',
			},
		])
		expect(wb.getWorkbookModel().chartParts[0]).toMatchObject({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'lineChart',
			title: 'Revenue',
			series: [
				{
					nameRef: 'Sheet1!$D$1',
					categoryRef: 'Sheet1!$A$2:$A$10',
					valueRef: 'Sheet1!$D$2:$D$10',
				},
				{
					nameText: 'Plan',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$C$2:$C$4',
				},
			],
		})

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.getWorkbookModel().chartParts[0]).toEqual({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			chartType: 'lineChart',
			title: 'Revenue',
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$B$2:$B$4',
				},
				{
					nameText: 'Plan',
					categoryRef: 'Sheet1!$A$2:$A$4',
					valueRef: 'Sheet1!$C$2:$C$4',
				},
			],
		})
	})

	test('chart series journals mark rollback lossy when refs cannot be unset', () => {
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().chartParts.push({
			partPath: 'xl/charts/chart1.xml',
			sheetName: 'Sheet1',
			series: [{ nameText: 'Actual', valueRef: 'Sheet1!$B$2:$B$4' }],
		})

		const journal = buildMutationJournal(wb.getWorkbookModel(), [
			{
				op: 'setChartSeriesSource',
				partPath: 'xl/charts/chart1.xml',
				seriesIndex: 0,
				nameRef: 'Sheet1!$B$1',
				valueRef: 'Sheet1!$C$2:$C$4',
			},
		])

		expect(journal.supported).toBe(true)
		expect(journal.exact).toBe(false)
		expect(journal.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Chart series selector cannot be restored exactly for series 0',
			},
		])
		expect(journal.inverseOps).toEqual([
			{
				op: 'setChartSeriesSource',
				partPath: 'xl/charts/chart1.xml',
				seriesIndex: 0,
				valueRef: 'Sheet1!$B$2:$B$4',
			},
		])
	})

	test('journal inverse ops restore pivot cache source and refresh metadata', () => {
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'Sheet1',
			name: 'PivotTable1',
			cacheId: 34,
			fields: [],
			rowFields: [],
			columnFields: [],
			pageFields: [],
			dataFields: [],
		})
		wb.getWorkbookModel().pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			relId: 'rIdPivotCache1',
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			recordsPartPath: 'xl/pivotCacheRecords/pivotCacheRecords1.xml',
			refreshOnLoad: false,
			enableRefresh: true,
			invalid: false,
			saveData: true,
			fields: [],
		})

		const changed = wb.apply(
			[
				{
					op: 'setPivotCache',
					pivotTable: 'PivotTable1',
					sourceSheet: 'RawData',
					sourceRef: 'A1:E20',
					refreshOnLoad: true,
					enableRefresh: false,
					invalid: true,
					saveData: false,
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setPivotCache',
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				sourceSheet: 'Raw',
				sourceRef: 'A1:D10',
				refreshOnLoad: false,
				enableRefresh: true,
				invalid: false,
				saveData: true,
			},
		])
		expect(wb.getWorkbookModel().pivotCaches[0]).toMatchObject({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			relId: 'rIdPivotCache1',
			recordsPartPath: 'xl/pivotCacheRecords/pivotCacheRecords1.xml',
			sourceSheet: 'RawData',
			sourceRef: 'A1:E20',
			refreshOnLoad: true,
			enableRefresh: false,
			invalid: true,
			saveData: false,
		})

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.getWorkbookModel().pivotCaches[0]).toEqual({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			relId: 'rIdPivotCache1',
			sourceSheet: 'Raw',
			sourceRef: 'A1:D10',
			recordsPartPath: 'xl/pivotCacheRecords/pivotCacheRecords1.xml',
			refreshOnLoad: false,
			enableRefresh: true,
			invalid: false,
			saveData: true,
			fields: [],
		})
	})

	test('pivot cache journals mark rollback lossy when public ops cannot unset fields', () => {
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			cacheId: 34,
			sourceRef: 'A1:D10',
			fields: [],
		})

		const journal = buildMutationJournal(wb.getWorkbookModel(), [
			{
				op: 'setPivotCache',
				cacheId: 34,
				sourceSheet: 'RawData',
				sourceRef: 'A1:E20',
				refreshOnLoad: true,
			},
		])

		expect(journal.supported).toBe(true)
		expect(journal.exact).toBe(false)
		expect(journal.issues).toEqual([
			{ code: 'LOSSY_INVERSE', message: 'Pivot cache selector cannot be restored exactly' },
		])
		expect(journal.inverseOps).toEqual([
			{
				op: 'setPivotCache',
				partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
				sourceRef: 'A1:D10',
			},
		])
	})

	test('journal inverse ops restore conditional format replacements', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B1:B10',
				rule: {
					type: 'cellIs',
					operator: 'greaterThan',
					formula: '10',
					priority: 3,
					style: { font: { bold: true } },
				},
			},
		])

		const changed = wb.apply(
			[
				{
					op: 'setConditionalFormat',
					sheet: 'Sheet1',
					range: 'B1:B10',
					rule: { type: 'expression', formula: 'B1<0', priority: 1 },
				},
			],
			{ journal: true },
		)

		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'B1:B10' },
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B1:B10',
				rule: {
					type: 'cellIs',
					operator: 'greaterThan',
					formula: '10',
					priority: 3,
					style: { font: { bold: true } },
				},
				mode: 'replace',
			},
		])

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.getConditionalFormats()).toEqual([
			{ type: 'cellIs', priority: 3, range: 'B1:B10' },
		])
	})

	test('journal inverse ops restore defined name edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A10' }])

		const changed = wb.apply([{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!B1:B10' }], {
			journal: true,
		})

		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A10' },
		])

		const undoChange = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undoChange.errors).toEqual([])
		expect(wb.definedName('Budget')?.formula).toBe('Sheet1!A1:A10')

		const deleted = wb.apply([{ op: 'deleteDefinedName', name: 'Budget' }], { journal: true })
		expect(deleted.journal?.inverseOps).toEqual([
			{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!A1:A10' },
		])

		const undoDelete = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDelete.errors).toEqual([])
		expect(wb.definedName('Budget')?.formula).toBe('Sheet1!A1:A10')

		const created = wb.apply([{ op: 'setDefinedName', name: 'Scratch', ref: 'Sheet1!C1:C2' }], {
			journal: true,
		})
		expect(created.journal?.inverseOps).toEqual([{ op: 'deleteDefinedName', name: 'Scratch' }])
	})

	test('journal inverse ops restore workbook and document properties', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setWorkbookProperties',
				properties: { codeName: 'OriginalBook', filterPrivacy: true, date1904: true },
			},
			{
				op: 'setDocumentProperties',
				properties: {
					core: { title: 'Original', creator: 'Analyst' },
					app: { Company: 'Ascend', TitlesOfParts: ['Sheet1'] },
					custom: [{ name: 'Desk', value: 'North', type: 'lpwstr', pid: 2 }],
				},
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{
					op: 'setWorkbookProperties',
					mode: 'replace',
					properties: { codeName: 'AgentBook', date1904: false },
				},
				{
					op: 'setDocumentProperties',
					mode: 'replace',
					properties: {
						core: { title: 'Changed', lastModifiedBy: 'agent' },
						app: { Company: 'Agent' },
						custom: [{ name: 'Risk', value: true, type: 'bool', pid: 3 }],
					},
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setDocumentProperties',
				mode: 'replace',
				properties: {
					core: { title: 'Original', creator: 'Analyst' },
					app: { Company: 'Ascend', TitlesOfParts: ['Sheet1'] },
					custom: [{ name: 'Desk', value: 'North', type: 'lpwstr', pid: 2 }],
				},
			},
			{
				op: 'setWorkbookProperties',
				mode: 'replace',
				properties: { codeName: 'OriginalBook', filterPrivacy: true, date1904: true },
			},
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal inverse ops restore workbook views and calc settings', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setWorkbookView', index: 0, view: { activeTab: 0, firstSheet: 0, tabRatio: 600 } },
			{
				op: 'setCalcSettings',
				settings: {
					calcMode: 'auto',
					fullCalcOnLoad: false,
					calcCompleted: true,
					calcOnSave: true,
					forceFullCalc: false,
					calcId: 1,
					dateSystem: '1900',
					iterativeCalc: { enabled: false, maxIterations: 100, maxChange: 0.001 },
				},
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{
					op: 'setWorkbookView',
					index: 0,
					mode: 'replace',
					view: { activeTab: 2, firstSheet: 1 },
				},
				{
					op: 'setWorkbookView',
					index: 1,
					view: { activeTab: 1, tabRatio: 720 },
				},
				{
					op: 'setCalcSettings',
					settings: {
						calcMode: 'manual',
						fullCalcOnLoad: true,
						calcCompleted: null,
						calcOnSave: false,
						forceFullCalc: true,
						calcId: 42,
						dateSystem: '1904',
						iterativeCalc: { enabled: true, maxIterations: 50, maxChange: 0.0001 },
					},
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setCalcSettings',
				settings: {
					calcMode: 'auto',
					fullCalcOnLoad: false,
					calcCompleted: true,
					calcOnSave: true,
					forceFullCalc: false,
					calcId: 1,
					dateSystem: '1900',
					iterativeCalc: { enabled: false, maxIterations: 100, maxChange: 0.001 },
				},
			},
			{ op: 'setWorkbookProperties', properties: { date1904: false }, mode: 'replace' },
			{ op: 'setWorkbookView', index: 1, view: null },
			{
				op: 'setWorkbookView',
				index: 0,
				view: { activeTab: 0, firstSheet: 0, tabRatio: 600 },
				mode: 'replace',
			},
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal inverse ops restore workbook protection when public ops can be exact', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setWorkbookProtection',
				protection: {
					lockStructure: true,
					workbookAlgorithmName: 'SHA-512',
					workbookSpinCount: 100000,
				},
			},
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{
					op: 'setWorkbookProtection',
					protection: { lockWindows: true, workbookPassword: 'ABCD' },
				},
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps).toEqual([
			{
				op: 'setWorkbookProtection',
				protection: {
					lockStructure: true,
					workbookAlgorithmName: 'SHA-512',
					workbookSpinCount: 100000,
				},
			},
		])
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)

		const unprotected = AscendWorkbook.create()
		const lossy = unprotected.preview(
			[{ op: 'setWorkbookProtection', protection: { lockStructure: true } }],
			{ journal: true },
		)
		expect(lossy.journal?.supported).toBe(true)
		expect(lossy.journal?.exact).toBe(false)
		expect(lossy.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Workbook protection absence cannot be restored exactly with public operations',
			},
		])
		expect(lossy.journal?.inverseOps).toEqual([{ op: 'setWorkbookProtection', protection: {} }])
	})

	test('journal exact inverse restores mixed semantic workbook state', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'C1', value: 'Amount' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 2 },
					{ ref: 'C2', value: 20 },
					{ ref: 'A3', value: 'East' },
					{ ref: 'B3', value: 3 },
					{ ref: 'C3', value: 30 },
				],
			},
			{ op: 'setStyle', sheet: 'Sheet1', range: 'B2:B2', style: { numberFormat: '0.00' } },
			{ op: 'setComment', sheet: 'Sheet1', ref: 'D4', text: 'initial', author: 'analyst' },
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'E5', url: 'https://example.com' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'A2:A5',
				rule: { type: 'list', formula1: '"West,East"', allowBlank: true },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'B2:B5',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '1', priority: 2 },
			},
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'F1:G1' },
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:C5', column: 0, values: ['West'] },
			{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$B$2:$B$3' },
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:C3', name: 'Sales', hasHeaders: true },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setTableColumn', table: 'Sales', column: 'Amount', formula: '=SUM([Qty])' },
			{ op: 'freezePane', sheet: 'Sheet1', row: 1, col: 1 },
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B2', value: 9 }] },
				{ op: 'setFormula', sheet: 'Sheet1', ref: 'H2', formula: 'B2*10' },
				{ op: 'setStyle', sheet: 'Sheet1', range: 'B2:B2', style: { numberFormat: '$0.00' } },
				{ op: 'setComment', sheet: 'Sheet1', ref: 'D4', text: 'changed', author: 'agent' },
				{
					op: 'setHyperlink',
					sheet: 'Sheet1',
					ref: 'E5',
					url: 'https://example.org',
					display: 'updated',
				},
				{
					op: 'setDataValidation',
					sheet: 'Sheet1',
					range: 'A2:A5',
					rule: { type: 'list', formula1: '"North,South"' },
				},
				{
					op: 'setConditionalFormat',
					sheet: 'Sheet1',
					range: 'B2:B5',
					rule: { type: 'expression', formula: 'B2>5', priority: 1 },
				},
				{ op: 'unmergeCells', sheet: 'Sheet1', range: 'F1:G1' },
				{ op: 'clearAutoFilter', sheet: 'Sheet1' },
				{ op: 'setDefinedName', name: 'Budget', ref: 'Sheet1!$C$2:$C$3' },
				{ op: 'renameTable', table: 'Sales', newName: 'Revenue' },
				{
					op: 'setTableColumn',
					table: 'Revenue',
					column: 'Amount',
					newName: 'RevenueAmount',
					formula: '=SUM([Qty])*2',
				},
				{ op: 'setTableStyle', table: 'Revenue', styleName: null },
				{ op: 'freezePane', sheet: 'Sheet1', row: 2, col: 0 },
				{ op: 'renameSheet', sheet: 'Sheet1', newName: 'Summary' },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(changed.journal?.inverseOps.length).toBeGreaterThan(0)
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal exact inverse restores deleted metadata surfaces', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setComment', sheet: 'Sheet1', ref: 'B2', text: 'review', author: 'analyst' },
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'C3',
				url: 'https://example.com',
				display: 'Report',
			},
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'D2:D10',
				rule: { type: 'whole', operator: 'greaterThan', formula1: '0', allowBlank: true },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'E2:E10',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '5', priority: 3 },
			},
			{ op: 'setAutoFilter', sheet: 'Sheet1', range: 'A1:E10', column: 0, values: ['Open'] },
			{ op: 'setDefinedName', name: 'ReviewRange', ref: 'Sheet1!$B$2:$E$10' },
		])
		const before = journalComparableState(wb)

		const changed = wb.apply(
			[
				{ op: 'deleteComment', sheet: 'Sheet1', ref: 'B2' },
				{ op: 'deleteHyperlink', sheet: 'Sheet1', ref: 'C3' },
				{ op: 'deleteDataValidation', sheet: 'Sheet1', range: 'D2:D10' },
				{ op: 'deleteConditionalFormat', sheet: 'Sheet1', range: 'E2:E10' },
				{ op: 'clearAutoFilter', sheet: 'Sheet1' },
				{ op: 'deleteDefinedName', name: 'ReviewRange' },
			],
			{ journal: true },
		)

		expect(changed.errors).toEqual([])
		expect(changed.journal?.supported).toBe(true)
		expect(changed.journal?.exact).toBe(true)
		expect(journalComparableState(wb)).not.toEqual(before)

		const undo = wb.apply(changed.journal?.inverseOps ?? [], { transaction: true })
		expect(undo.errors).toEqual([])
		expect(journalComparableState(wb)).toEqual(before)
	})

	test('journal inverse ops restore table renames and column metadata edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Region' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'C1', value: 'Amount' },
					{ ref: 'A2', value: 'West' },
					{ ref: 'B2', value: 2 },
					{ ref: 'C2', value: 5 },
					{ ref: 'A3', value: 'Total' },
					{ ref: 'B3', value: 2 },
					{ ref: 'C3', value: 5 },
				],
			},
			{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:C3', name: 'Sales', hasHeaders: true },
			{
				op: 'setTableColumn',
				table: 'Sales',
				column: 'Amount',
				formula: '=[@Qty]*10',
				totalsRowFunction: 'sum',
			},
		])
		const internal = wb as unknown as {
			wb: {
				sheets: Array<{
					tables: Array<{
						hasTotals?: boolean
						ref: { start: { row: number; col: number }; end: { row: number; col: number } }
					}>
				}>
			}
		}
		const tableModel = internal.wb.sheets[0]?.tables[0]
		if (tableModel) {
			tableModel.hasTotals = true
			tableModel.ref = { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } }
		}

		const columnEdit = wb.apply(
			[
				{
					op: 'setTableColumn',
					table: 'Sales',
					column: 'Amount',
					newName: 'Revenue',
					formula: '=[@Qty]*12',
					totalsRowFunction: 'average',
				},
			],
			{ journal: true },
		)

		expect(columnEdit.journal?.exact).toBe(true)
		expect(columnEdit.journal?.inverseOps).toEqual([
			{
				op: 'setTableColumn',
				table: 'Sales',
				column: 'Revenue',
				newName: 'Amount',
				formula: '[@Qty]*10',
				totalsRowFunction: 'sum',
			},
		])

		const undoColumn = wb.apply(columnEdit.journal?.inverseOps ?? [], { transaction: true })
		expect(undoColumn.errors).toEqual([])
		expect(wb.table('Sales')?.columns).toEqual(['Region', 'Qty', 'Amount'])
		expect(wb.table('Sales')?.columnDefs[2]).toMatchObject({
			name: 'Amount',
			formula: '[@Qty]*10',
			totalsRowFunction: 'sum',
		})

		const renamed = wb.apply([{ op: 'renameTable', table: 'Sales', newName: 'RevenueTable' }], {
			journal: true,
		})
		expect(renamed.journal?.inverseOps).toEqual([
			{ op: 'renameTable', table: 'RevenueTable', newName: 'Sales' },
		])

		const undoRename = wb.apply(renamed.journal?.inverseOps ?? [], { transaction: true })
		expect(undoRename.errors).toEqual([])
		expect(wb.table('Sales')?.name).toBe('Sales')
	})

	test('journal inverse ops restore table lifecycle metadata edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'Product' },
					{ ref: 'B1', value: 'Qty' },
					{ ref: 'A2', value: 'Widget' },
					{ ref: 'B2', value: 2 },
					{ ref: 'A3', value: 'Gadget' },
					{ ref: 'B3', value: 3 },
				],
			},
		])

		const created = wb.apply(
			[{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B3', name: 'Sales', hasHeaders: true }],
			{ journal: true },
		)
		expect(created.journal?.exact).toBe(true)
		expect(created.journal?.inverseOps).toEqual([{ op: 'deleteTable', table: 'Sales' }])

		wb.apply([
			{ op: 'setTableColumn', table: 'Sales', column: 'Qty', formula: '=1+1' },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
		])

		const resized = wb.apply([{ op: 'resizeTable', table: 'Sales', ref: 'A1:B2' }], {
			journal: true,
		})
		expect(resized.journal?.exact).toBe(true)
		expect(resized.journal?.inverseOps).toEqual([
			{ op: 'resizeTable', table: 'Sales', ref: 'A1:B3' },
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setTableColumn', table: 'Sales', column: 1, formula: '1+1' },
		])

		const undoResize = wb.apply(resized.journal?.inverseOps ?? [], { transaction: true })
		expect(undoResize.errors).toEqual([])
		expect(wb.table('Sales')?.ref).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 2, col: 1 },
		})
		expect(wb.table('Sales')?.columnDefs[1]?.formula).toBe('1+1')

		const deleted = wb.apply([{ op: 'deleteTable', table: 'Sales' }], { journal: true })
		expect(deleted.journal?.exact).toBe(true)
		expect(deleted.journal?.inverseOps).toEqual([
			{
				op: 'createTable',
				sheet: 'Sheet1',
				ref: 'A1:B3',
				name: 'Sales',
				hasHeaders: true,
			},
			{ op: 'setTableStyle', table: 'Sales', styleName: 'TableStyleMedium2' },
			{ op: 'setTableColumn', table: 'Sales', column: 0, newName: 'Product' },
			{ op: 'setTableColumn', table: 'Sales', column: 1, newName: 'Qty', formula: '1+1' },
		])

		const undoDelete = wb.apply(deleted.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDelete.errors).toEqual([])
		expect(wb.table('Sales')?.columns).toEqual(['Product', 'Qty'])
		expect(wb.table('Sales')?.columnDefs[1]?.formula).toBe('1+1')
		expect(wb.table('Sales')?.styleInfo).toEqual({ name: 'TableStyleMedium2' })
	})

	test('journal inverse ops restore structural row and column edits', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 'top' },
					{ ref: 'A2', value: 'deleted-row' },
					{ ref: 'B2', value: 4 },
					{ ref: 'A3', value: 'bottom' },
				],
			},
			{ op: 'setStyle', sheet: 'Sheet1', range: 'B2:B2', style: { numberFormat: '0.00' } },
		])

		const inserted = wb.apply([{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})
		expect(inserted.journal?.exact).toBe(true)
		expect(inserted.journal?.inverseOps).toEqual([
			{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 },
		])
		const undoInsert = wb.apply(inserted.journal?.inverseOps ?? [], { transaction: true })
		expect(undoInsert.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({
			kind: 'string',
			value: 'deleted-row',
		})

		const deletedRow = wb.apply([{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})
		expect(deletedRow.journal?.exact).toBe(true)
		expect(deletedRow.journal?.inverseOps).toEqual([
			{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 },
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A2', value: 'deleted-row' },
					{ ref: 'B2', value: 4 },
				],
			},
			{ op: 'setStyle', sheet: 'Sheet1', range: 'A2', style: {} },
			{ op: 'setStyle', sheet: 'Sheet1', range: 'B2', style: { numberFormat: '0.00' } },
		])
		const undoDeleteRow = wb.apply(deletedRow.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDeleteRow.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('A2')?.value).toEqual({
			kind: 'string',
			value: 'deleted-row',
		})
		expect(wb.cellStyle('Sheet1!B2')?.numberFormat).toBe('0.00')

		const deletedCol = wb.apply([{ op: 'deleteCols', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})
		expect(deletedCol.journal?.exact).toBe(true)
		expect(deletedCol.journal?.inverseOps[0]).toEqual({
			op: 'insertCols',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
		})
		const undoDeleteCol = wb.apply(deletedCol.journal?.inverseOps ?? [], { transaction: true })
		expect(undoDeleteCol.errors).toEqual([])
		expect(wb.sheet('Sheet1')?.cell('B2')?.value).toEqual({ kind: 'number', value: 4 })
	})

	test('structural delete journals mark metadata intersections as lossy', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'header' }] },
			{ op: 'mergeCells', sheet: 'Sheet1', range: 'C2:D2' },
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'F2:F2',
				rule: { type: 'list', formula1: '"yes,no"' },
			},
		])

		const deletedRow = wb.preview([{ op: 'deleteRows', sheet: 'Sheet1', at: 1, count: 1 }], {
			journal: true,
		})

		expect(deletedRow.wouldSucceed).toBe(true)
		expect(deletedRow.journal?.supported).toBe(true)
		expect(deletedRow.journal?.exact).toBe(false)
		expect(deletedRow.journal?.issues).toEqual([
			{
				code: 'LOSSY_INVERSE',
				message: 'Deleted row metadata on Sheet1 cannot be fully restored with public operations',
				refs: ['Sheet1!2'],
			},
		])
		expect(deletedRow.journal?.inverseOps[0]).toEqual({
			op: 'insertRows',
			sheet: 'Sheet1',
			at: 1,
			count: 1,
		})
		expect(wb.sheet('Sheet1')?.merges).toEqual([
			{ start: { row: 1, col: 2 }, end: { row: 1, col: 3 } },
		])
	})
})

function journalComparableState(wb: AscendWorkbook): object {
	const snapshot = wb.snapshot()
	return {
		sheets: snapshot.sheets,
		names: snapshot.names,
		workbookProtectionJson: snapshot.workbookProtectionJson,
		workbookProperties: wb.getWorkbookModel().workbookProperties,
		documentProperties: wb.inspect().documentProperties,
		workbookViews: wb.inspect().workbookViews,
		calcSettings: wb.getWorkbookModel().calcSettings,
		sheetMetadata: wb.sheets.map((name) => {
			const sheet = wb.sheet(name)
			return {
				name,
				frozenRows: sheet?.frozenRows,
				frozenCols: sheet?.frozenCols,
				merges: sheet?.merges,
				dataValidations: sheet?.dataValidations,
				conditionalFormats: sheet?.conditionalFormats,
				comments: sheet?.getComments(),
				hyperlinks: sheet?.getHyperlinks(),
				autoFilter: sheet?.autoFilter,
			}
		}),
		styles: {
			b2: wb.cellStyle(`${wb.sheets[0]}!B2`),
			c2: wb.cellStyle(`${wb.sheets[0]}!C2`),
		},
	}
}
