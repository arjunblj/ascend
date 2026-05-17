import { describe, expect, test } from 'bun:test'
import { createTableId, createWorkbook } from './index.ts'

describe('Workbook.clone', () => {
	test('clones sheet table metadata without aliasing nested refs', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		sheet.autoFilter = { ref: 'A1:B3', columns: [], sortState: { ref: 'A2:B3', conditions: [] } }
		sheet.preservedSheetViewSelections = [{ pane: 'bottomRight', activeCell: 'B2', sqref: 'B2' }]
		sheet.preservedCellMetadata.set('0:0', { cm: 1 })
		sheet.pageSetupPr = { fitToPage: true, autoPageBreaks: false }
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
		;(cloneSheet.preservedSheetViewSelections?.[0] as { activeCell: string }).activeCell = 'C3'
		;(cloneSheet.preservedCellMetadata.get('0:0') as { cm: number }).cm = 2
		;(cloneSheet.pageSetupPr as { fitToPage: boolean }).fitToPage = false

		expect(sheet.tables[0]?.ref.start.row).toBe(0)
		expect(sheet.tables[0]?.columns[0]?.name).toBe('Name')
		expect(sheet.autoFilter?.ref).toBe('A1:B3')
		expect(sheet.preservedSheetViewSelections?.[0]?.activeCell).toBe('B2')
		expect(sheet.preservedCellMetadata.get('0:0')?.cm).toBe(1)
		expect(sheet.pageSetupPr?.fitToPage).toBe(true)
	})

	test('clones workbook settings and preserved metadata without aliasing', () => {
		const wb = createWorkbook()
		wb.calcSettings = {
			...wb.calcSettings,
			iterativeCalc: { enabled: true, maxIterations: 10, maxChange: 0.1 },
			extraAttributes: [{ name: 'fullPrecision', value: '0' }],
		}
		wb.preservedStyles = {
			xfByStyleId: { 0: 1 },
			baseStyleIdByStyleId: { 0: 0 },
		}
		wb.workbookProperties = {
			codeName: 'Model',
			extraAttributes: [{ name: 'checkCompatibility', value: '1' }],
		}
		wb.workbookProtection = {
			lockStructure: true,
			extraAttributes: [{ name: 'futureProtectionMode', value: 'strict' }],
		}
		wb.workbookViews.push({
			activeTab: 0,
			extraAttributes: [{ name: 'windowWidth', value: '16800' }],
		})

		const clone = wb.clone()
		;(clone.calcSettings.iterativeCalc as { enabled: boolean }).enabled = false
		;(clone.calcSettings.extraAttributes?.[0] as { value: string }).value = '1'
		if (clone.preservedStyles) {
			;(clone.preservedStyles.xfByStyleId as Record<number, number>)[0] = 99
		}
		;(clone.workbookProperties.extraAttributes?.[0] as { value: string }).value = '0'
		;(clone.workbookProtection?.extraAttributes?.[0] as { value: string }).value = 'legacy'
		;(clone.workbookViews[0]?.extraAttributes?.[0] as { value: string }).value = '9000'

		expect(wb.calcSettings.iterativeCalc.enabled).toBe(true)
		expect(wb.calcSettings.extraAttributes?.[0]?.value).toBe('0')
		expect(wb.preservedStyles?.xfByStyleId[0]).toBe(1)
		expect(wb.workbookProperties.extraAttributes?.[0]?.value).toBe('1')
		expect(wb.workbookProtection?.extraAttributes?.[0]?.value).toBe('strict')
		expect(wb.workbookViews[0]?.extraAttributes?.[0]?.value).toBe('16800')
	})

	test('clones document properties without aliasing nested collections', () => {
		const wb = createWorkbook()
		wb.documentProperties = {
			core: { creator: 'Analyst' },
			app: { Application: 'Excel', TitlesOfParts: ['Sheet1'] },
			custom: [{ name: 'Desk', value: 'Research', type: 'lpwstr', pid: 2 }],
		}

		const clone = wb.clone()
		;(clone.documentProperties.core as { creator: string }).creator = 'Reviewer'
		;(clone.documentProperties.app?.TitlesOfParts as string[])[0] = 'Changed'
		;(clone.documentProperties.custom?.[0] as { name: string }).name = 'Changed'

		expect(wb.documentProperties.core?.creator).toBe('Analyst')
		expect(wb.documentProperties.app?.TitlesOfParts).toEqual(['Sheet1'])
		expect(wb.documentProperties.custom?.[0]?.name).toBe('Desk')
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
			records: {
				partPath: 'xl/pivotCache/pivotCacheRecords1.xml',
				declaredCount: 1,
				parsedCount: 1,
				preview: [
					{
						index: 0,
						values: [{ index: 0, kind: 'sharedItem', sharedItemIndex: 0 }],
					},
				],
				materializedRecords: [
					{
						index: 0,
						values: [{ index: 0, kind: 'sharedItem', sharedItemIndex: 0 }],
					},
				],
				materializedCount: 1,
				materializedComplete: true,
				valueKindCounts: [{ kind: 'sharedItem', count: 1 }],
			},
			fields: [
				{
					index: 0,
					name: 'Region',
					sharedItemsInfo: { count: 1, containsString: true },
					sharedItems: [{ index: 0, kind: 'string', value: 'West' }],
					fieldGroup: {
						base: 2,
						range: { groupBy: 'months', startDate: '2024-01-01T00:00:00' },
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
			rowItems: [{ index: 0, fieldItems: [{ index: 0, item: 1 }] }],
			columnItems: [{ index: 0, dataFieldIndex: 1, fieldItems: [{ index: 0, item: 2 }] }],
			formats: [
				{
					index: 0,
					dxfId: 3,
					area: {
						fieldPosition: 0,
						references: [{ index: 0, field: 1, itemCount: 1, items: [{ index: 0, item: 2 }] }],
					},
				},
			],
			chartFormats: [
				{
					index: 0,
					chart: 2,
					formatId: 3,
					series: true,
					area: {
						fieldPosition: 0,
						references: [{ index: 0, field: 1, itemCount: 1, items: [{ index: 0, item: 2 }] }],
					},
				},
			],
		})

		const clone = wb.clone()
		const pivotLocation = clone.pivotTables[0]?.location
		const pivotOptions = clone.pivotTables[0]?.options
		const pivotStyle = clone.pivotTables[0]?.style
		const sharedItemsInfo = clone.pivotCaches[0]?.fields[0]?.sharedItemsInfo
		const sharedItem = clone.pivotCaches[0]?.fields[0]?.sharedItems?.[0]
		const cacheRecord = clone.pivotCaches[0]?.records?.preview[0]
		const cacheRecordValue = cacheRecord?.values[0]
		const materializedRecord = clone.pivotCaches[0]?.records?.materializedRecords?.[0]
		const materializedRecordValue = materializedRecord?.values[0]
		const cacheRecordKindCount = clone.pivotCaches[0]?.records?.valueKindCounts[0]
		const groupRange = clone.pivotCaches[0]?.fields[0]?.fieldGroup?.range
		const discreteItem = clone.pivotCaches[0]?.fields[0]?.fieldGroup?.discreteItems?.[0]
		const groupItem = clone.pivotCaches[0]?.fields[0]?.fieldGroup?.groupItems?.[0]
		const item = clone.pivotTables[0]?.fields[0]?.items?.[0]
		const rowItem = clone.pivotTables[0]?.rowItems?.[0]?.fieldItems[0]
		const columnItem = clone.pivotTables[0]?.columnItems?.[0]?.fieldItems[0]
		const format = clone.pivotTables[0]?.formats?.[0]
		const formatReference = format?.area?.references?.[0]
		const formatReferenceItem = formatReference?.items[0]
		const chartFormat = clone.pivotTables[0]?.chartFormats?.[0]
		const chartFormatReference = chartFormat?.area?.references?.[0]
		const chartFormatReferenceItem = chartFormatReference?.items[0]
		expect(pivotLocation).toBeDefined()
		expect(pivotOptions).toBeDefined()
		expect(pivotStyle).toBeDefined()
		expect(sharedItemsInfo).toBeDefined()
		expect(sharedItem).toBeDefined()
		expect(cacheRecord).toBeDefined()
		expect(cacheRecordValue).toBeDefined()
		expect(materializedRecord).toBeDefined()
		expect(materializedRecordValue).toBeDefined()
		expect(cacheRecordKindCount).toBeDefined()
		expect(groupRange).toBeDefined()
		expect(discreteItem).toBeDefined()
		expect(groupItem).toBeDefined()
		expect(item).toBeDefined()
		expect(rowItem).toBeDefined()
		expect(columnItem).toBeDefined()
		expect(format).toBeDefined()
		expect(formatReference).toBeDefined()
		expect(formatReferenceItem).toBeDefined()
		expect(chartFormat).toBeDefined()
		expect(chartFormatReference).toBeDefined()
		expect(chartFormatReferenceItem).toBeDefined()
		if (
			!pivotLocation ||
			!pivotOptions ||
			!pivotStyle ||
			!sharedItemsInfo ||
			!sharedItem ||
			!cacheRecord ||
			!cacheRecordValue ||
			!materializedRecord ||
			!materializedRecordValue ||
			!cacheRecordKindCount ||
			!groupRange ||
			!discreteItem ||
			!groupItem ||
			!item ||
			!rowItem ||
			!columnItem ||
			!format ||
			!formatReference ||
			!formatReferenceItem ||
			!chartFormat ||
			!chartFormatReference ||
			!chartFormatReferenceItem
		) {
			return
		}

		;(pivotLocation as { ref: string }).ref = 'B4:E21'
		;(pivotOptions as { updatedVersion: number }).updatedVersion = 8
		;(pivotStyle as { name: string }).name = 'PivotStyleDark1'
		;(sharedItemsInfo as { count: number }).count = 2
		;(sharedItem as { value: string }).value = 'East'
		;(cacheRecord as { index: number }).index = 1
		;(cacheRecordValue as { sharedItemIndex: number }).sharedItemIndex = 2
		;(materializedRecord as { index: number }).index = 1
		;(materializedRecordValue as { sharedItemIndex: number }).sharedItemIndex = 2
		;(cacheRecordKindCount as { count: number }).count = 2
		;(groupRange as { groupBy: string }).groupBy = 'quarters'
		;(discreteItem as { value: number }).value = 0
		;(groupItem as { value: string }).value = 'Group2'
		;(item as { hidden: boolean }).hidden = false
		;(rowItem as { item: number }).item = 4
		;(columnItem as { item: number }).item = 5
		;(format as { dxfId: number }).dxfId = 4
		;(formatReference as { itemCount: number }).itemCount = 2
		;(formatReferenceItem as { item: number }).item = 6
		;(chartFormat as { formatId: number }).formatId = 4
		;(chartFormatReference as { itemCount: number }).itemCount = 2
		;(chartFormatReferenceItem as { item: number }).item = 6

		expect(wb.pivotTables[0]?.location?.ref).toBe('A3:D20')
		expect(wb.pivotTables[0]?.options?.updatedVersion).toBe(7)
		expect(wb.pivotTables[0]?.style?.name).toBe('PivotStyleLight16')
		expect(wb.pivotTables[0]?.rowItems?.[0]?.fieldItems[0]?.item).toBe(1)
		expect(wb.pivotTables[0]?.columnItems?.[0]?.fieldItems[0]?.item).toBe(2)
		expect(wb.pivotTables[0]?.formats?.[0]?.dxfId).toBe(3)
		expect(wb.pivotTables[0]?.formats?.[0]?.area?.references?.[0]?.itemCount).toBe(1)
		expect(wb.pivotTables[0]?.formats?.[0]?.area?.references?.[0]?.items[0]?.item).toBe(2)
		expect(wb.pivotTables[0]?.chartFormats?.[0]?.formatId).toBe(3)
		expect(wb.pivotTables[0]?.chartFormats?.[0]?.area?.references?.[0]?.itemCount).toBe(1)
		expect(wb.pivotTables[0]?.chartFormats?.[0]?.area?.references?.[0]?.items[0]?.item).toBe(2)
		expect(wb.pivotCaches[0]?.fields[0]?.sharedItemsInfo?.count).toBe(1)
		expect(wb.pivotCaches[0]?.fields[0]?.sharedItems?.[0]?.value).toBe('West')
		expect(wb.pivotCaches[0]?.records?.preview[0]?.index).toBe(0)
		expect(wb.pivotCaches[0]?.records?.preview[0]?.values[0]?.sharedItemIndex).toBe(0)
		expect(wb.pivotCaches[0]?.records?.materializedRecords?.[0]?.index).toBe(0)
		expect(wb.pivotCaches[0]?.records?.materializedRecords?.[0]?.values[0]?.sharedItemIndex).toBe(0)
		expect(wb.pivotCaches[0]?.records?.valueKindCounts[0]?.count).toBe(1)
		expect(wb.pivotCaches[0]?.fields[0]?.fieldGroup?.range?.groupBy).toBe('months')
		expect(wb.pivotCaches[0]?.fields[0]?.fieldGroup?.discreteItems?.[0]?.value).toBe(1)
		expect(wb.pivotCaches[0]?.fields[0]?.fieldGroup?.groupItems?.[0]?.value).toBe('Group1')
		expect(wb.pivotTables[0]?.fields[0]?.items?.[0]?.hidden).toBe(true)
	})
})
