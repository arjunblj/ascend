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

	test('clones pivot field item inventory without aliasing', () => {
		const wb = createWorkbook()
		wb.pivotCaches.push({
			partPath: 'xl/pivotCache/pivotCacheDefinition1.xml',
			fields: [
				{
					index: 0,
					name: 'Region',
					sharedItemsInfo: { count: 1, containsString: true },
					sharedItems: [{ index: 0, kind: 'string', value: 'West' }],
					fieldGroup: {
						base: 2,
						discreteItems: [{ index: 0, value: 1 }],
						groupItems: [{ index: 0, kind: 'string', value: 'Group1' }],
					},
				},
			],
		})
		wb.pivotTables.push({
			partPath: 'xl/pivotTables/pivotTable1.xml',
			sheetName: 'PivotSheet',
			name: 'PivotTable1',
			location: { ref: 'A3:D20', firstDataRow: 1 },
			options: { dataOnRows: true, updatedVersion: 7 },
			style: { name: 'PivotStyleLight16', showRowHeaders: true },
			fields: [
				{
					index: 0,
					axis: 'axisPage',
					items: [{ index: 0, cacheIndex: 1, hidden: true }],
				},
			],
			rowFields: [],
			columnFields: [],
			pageFields: [{ index: 0, item: 1 }],
			dataFields: [{ fieldIndex: 1, showDataAs: 'percent', baseField: 0, baseItem: 2 }],
		})

		const clone = wb.clone()
		const pivotLocation = clone.pivotTables[0]?.location
		const pivotOptions = clone.pivotTables[0]?.options
		const pivotStyle = clone.pivotTables[0]?.style
		const sharedItemsInfo = clone.pivotCaches[0]?.fields[0]?.sharedItemsInfo
		const sharedItem = clone.pivotCaches[0]?.fields[0]?.sharedItems?.[0]
		const discreteItem = clone.pivotCaches[0]?.fields[0]?.fieldGroup?.discreteItems?.[0]
		const groupItem = clone.pivotCaches[0]?.fields[0]?.fieldGroup?.groupItems?.[0]
		const item = clone.pivotTables[0]?.fields[0]?.items?.[0]
		expect(pivotLocation).toBeDefined()
		expect(pivotOptions).toBeDefined()
		expect(pivotStyle).toBeDefined()
		expect(sharedItemsInfo).toBeDefined()
		expect(sharedItem).toBeDefined()
		expect(discreteItem).toBeDefined()
		expect(groupItem).toBeDefined()
		expect(item).toBeDefined()
		if (
			!pivotLocation ||
			!pivotOptions ||
			!pivotStyle ||
			!sharedItemsInfo ||
			!sharedItem ||
			!discreteItem ||
			!groupItem ||
			!item
		) {
			return
		}

		;(pivotLocation as { ref: string }).ref = 'B4:E21'
		;(pivotOptions as { updatedVersion: number }).updatedVersion = 8
		;(pivotStyle as { name: string }).name = 'PivotStyleDark1'
		;(sharedItemsInfo as { count: number }).count = 2
		;(sharedItem as { value: string }).value = 'East'
		;(discreteItem as { value: number }).value = 0
		;(groupItem as { value: string }).value = 'Group2'
		;(item as { hidden: boolean }).hidden = false

		expect(wb.pivotTables[0]?.location?.ref).toBe('A3:D20')
		expect(wb.pivotTables[0]?.options?.updatedVersion).toBe(7)
		expect(wb.pivotTables[0]?.style?.name).toBe('PivotStyleLight16')
		expect(wb.pivotCaches[0]?.fields[0]?.sharedItemsInfo?.count).toBe(1)
		expect(wb.pivotCaches[0]?.fields[0]?.sharedItems?.[0]?.value).toBe('West')
		expect(wb.pivotCaches[0]?.fields[0]?.fieldGroup?.discreteItems?.[0]?.value).toBe(1)
		expect(wb.pivotCaches[0]?.fields[0]?.fieldGroup?.groupItems?.[0]?.value).toBe('Group1')
		expect(wb.pivotTables[0]?.fields[0]?.items?.[0]?.hidden).toBe(true)
	})
})
