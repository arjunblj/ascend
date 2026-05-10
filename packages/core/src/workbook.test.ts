import { describe, expect, test } from 'bun:test'
import { createTableId, createWorkbook } from './index.ts'

describe('Workbook.clone', () => {
	test('clones sheet table metadata without aliasing nested refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = { ref: 'A1:B3', columns: [], sortState: { ref: 'A2:B3', conditions: [] } }
		sheet.tables.push({
			id: createTableId(),
			name: 'Data',
			sheetId: sheet.id,
			ref: { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } },
			columns: [{ name: 'Name' }, { name: 'Value' }],
			hasHeaders: true,
			hasTotals: false,
			autoFilter: { ref: 'A1:B3', columns: [] },
		})

		const clone = wb.clone()
		const cloneSheet = clone.getSheet('Sheet1')
		expect(cloneSheet).toBeDefined()
		if (!cloneSheet) return

		cloneSheet.ensureWritable()

		const cloneTable = cloneSheet.tables[0]
		expect(cloneTable).toBeDefined()
		if (!cloneTable) return

		;(cloneTable.ref.start as { row: number }).row = 10
		;(cloneTable.columns[0] as { name: string }).name = 'Changed'
		;(cloneSheet.autoFilter as { ref: string }).ref = 'C1:D3'

		expect(sheet.tables[0]?.ref.start.row).toBe(0)
		expect(sheet.tables[0]?.columns[0]?.name).toBe('Name')
		expect(sheet.autoFilter?.ref).toBe('A1:B3')
	})

	test('clones workbook settings and preserved metadata without aliasing', () => {
		const wb = createWorkbook()
		wb.calcSettings = {
			...wb.calcSettings,
			iterativeCalc: { enabled: true, maxIterations: 10, maxChange: 0.1 },
		}
		wb.preservedStyles = {
			xfByStyleId: { 0: 1 },
			baseStyleIdByStyleId: { 0: 0 },
		}

		const clone = wb.clone()
		;(clone.calcSettings.iterativeCalc as { enabled: boolean }).enabled = false
		if (clone.preservedStyles) {
			;(clone.preservedStyles.xfByStyleId as Record<number, number>)[0] = 99
		}

		expect(wb.calcSettings.iterativeCalc.enabled).toBe(true)
		expect(wb.preservedStyles?.xfByStyleId[0]).toBe(1)
	})

	test('clones active content VBA summaries without aliasing nested modules', () => {
		const wb = createWorkbook()
		wb.activeContent.push({
			kind: 'vbaProject',
			partPath: 'xl/vbaProject.bin',
			contentType: 'application/vnd.ms-office.vbaProject',
			anchor: 'workbook',
			relationshipCount: 0,
			opaque: true,
			executionPolicy: 'blocked',
			vbaProject: {
				moduleCount: 1,
				projectStreamPresent: true,
				modules: [{ name: 'Module1', kind: 'standard' }],
			},
		})

		const clone = wb.clone()
		const module = clone.activeContent[0]?.vbaProject?.modules[0]
		expect(module).toBeDefined()
		if (!module) return

		;(module as { name: string }).name = 'Changed'

		expect(wb.activeContent[0]?.vbaProject?.modules[0]?.name).toBe('Module1')
	})

	test('clones macro sheet inventory without aliasing', () => {
		const wb = createWorkbook()
		wb.macroSheets.push({
			name: 'Macro1',
			sheetId: '2',
			relId: 'rIdMacro',
			partPath: 'xl/macrosheets/sheet1.xml',
			state: 'veryHidden',
			relationshipCount: 0,
			dimensionRef: 'A1',
			cellCount: 1,
			formulaCount: 1,
		})

		const clone = wb.clone()
		const macroSheet = clone.macroSheets[0]
		expect(macroSheet).toBeDefined()
		if (!macroSheet) return

		;(macroSheet as { name: string }).name = 'Changed'

		expect(wb.macroSheets[0]?.name).toBe('Macro1')
	})
})
