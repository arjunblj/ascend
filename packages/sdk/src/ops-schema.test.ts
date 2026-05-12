import { describe, expect, test } from 'bun:test'
import { getOperationsSchema, listOperations, parseOperations } from './ops.ts'

describe('operation schema agent DX', () => {
	test('schemas include standard-schema-friendly metadata and recovery guidance', () => {
		const schemas = getOperationsSchema()
		const setCells = schemas.find((schema) => schema.op === 'setCells')
		expect(setCells?.schemaDialect).toBe('json-schema-draft-2020-12-compatible')
		expect(setCells?.standardSchema).toMatchObject({
			version: 1,
			vendor: 'ascend',
			name: 'setCells',
		})
		expect(setCells?.examples.length).toBeGreaterThan(0)
		expect(setCells?.invalidExamples[0]?.recoveryAction).toContain('required fields')
		expect(setCells?.recoveryActions.join('\n')).toContain('updates')
	})

	test('destructive operations advertise approval requirements', () => {
		const deleteSheet = listOperations().find((operation) => operation.op === 'deleteSheet')
		const deleteSheetSchema = getOperationsSchema().find((schema) => schema.op === 'deleteSheet')
		expect(deleteSheet?.approval?.required).toBe(true)
		expect(deleteSheetSchema?.approval?.approvalHint).toContain('approval id')
	})

	test('parse failures remain concise while schemas provide recovery context', () => {
		const parsed = parseOperations([{ op: 'setCells', sheet: 'Sheet1' }])
		expect(parsed.ok).toBe(false)
		if (!parsed.ok) {
			expect(parsed.issues[0]).toContain('updates is required')
		}
	})

	test('parse failures reject wrong field types before operation execution', () => {
		const parsed = parseOperations([
			{ op: 'insertRows', sheet: 'Sheet1', at: 0, count: '2' },
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: { nested: true } }] },
			{ op: 'copyRange', sheet: 'Sheet1', source: 'A1', target: 'B1', mode: 'everything' },
		])
		expect(parsed.ok).toBe(false)
		if (!parsed.ok) {
			expect(parsed.issues).toContain('ops[0].count must be a positive integer')
			expect(parsed.issues).toContain('ops[1].updates[0].value must be a scalar value or null')
			expect(parsed.issues[2]).toContain('ops[2].mode must be one of')
		}
	})

	test('parse failures reject operation-specific unknown fields', () => {
		const parsed = parseOperations([
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'A1', formula: '=1', count: 1 },
		])
		expect(parsed.ok).toBe(false)
		if (!parsed.ok) {
			expect(parsed.issues[0]).toContain('ops[0].count is not valid for setFormula')
		}
	})

	test('setHyperlink supports external urls, internal locations, and tooltip metadata', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setHyperlink')
		expect(schema?.schema.required).toEqual(['op', 'sheet', 'ref'])
		expect(schema?.schema.properties.location?.description).toContain('Internal workbook')
		expect(schema?.schema.properties.tooltip?.description).toContain('Tooltip')

		const parsed = parseOperations([
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', url: 'https://example.com' },
			{
				op: 'setHyperlink',
				sheet: 'Sheet1',
				ref: 'B1',
				location: 'Sheet2!A1',
				display: 'Jump',
				tooltip: 'Open Sheet2',
			},
		])
		expect(parsed.ok).toBe(true)

		const missingDestination = parseOperations([
			{ op: 'setHyperlink', sheet: 'Sheet1', ref: 'A1', display: 'No target' },
		])
		expect(missingDestination.ok).toBe(false)
		if (!missingDestination.ok) {
			expect(missingDestination.issues[0]).toContain('url or ops[0].location is required')
		}
	})

	test('copyRange and moveRange expose optional cross-sheet destination', () => {
		const copySchema = getOperationsSchema().find((entry) => entry.op === 'copyRange')
		const moveSchema = getOperationsSchema().find((entry) => entry.op === 'moveRange')
		expect(copySchema?.schema.properties.targetSheet?.description).toContain('Destination sheet')
		expect(moveSchema?.schema.properties.targetSheet?.type).toBe('string')

		const parsed = parseOperations([
			{
				op: 'copyRange',
				sheet: 'Sheet1',
				source: 'A1:B2',
				targetSheet: 'Summary',
				target: 'C3',
				mode: 'all',
			},
		])
		expect(parsed.ok).toBe(true)
	})

	test('replaceImage is exposed with selector guidance for visual edits', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'replaceImage')
		expect(schema?.schema.required).toEqual(['op', 'sheet', 'contentBase64', 'contentType'])
		expect(schema?.schema.properties.targetPath?.description).toContain('xl/media')
		expect(schema?.examples[0]).toMatchObject({
			op: 'replaceImage',
			sheet: 'Sheet1',
			contentType: 'image/png',
		})
		expect(schema?.recoveryActions.join('\n')).toContain('imageIndex')
	})

	test('insertImage and deleteImage expose image lifecycle guidance', () => {
		const insert = getOperationsSchema().find((entry) => entry.op === 'insertImage')
		expect(insert?.schema.required).toEqual(['op', 'sheet', 'contentBase64', 'contentType'])
		expect(insert?.schema.properties.anchor?.description).toContain('Image anchor')
		expect(insert?.examples[0]).toMatchObject({
			op: 'insertImage',
			sheet: 'Sheet1',
			contentType: 'image/png',
			name: 'Logo',
		})

		const deleteImage = getOperationsSchema().find((entry) => entry.op === 'deleteImage')
		expect(deleteImage?.schema.required).toEqual(['op', 'sheet'])
		expect(deleteImage?.examples[0]).toMatchObject({ op: 'deleteImage', imageIndex: 0 })
		expect(deleteImage?.recoveryActions.join('\n')).toContain('imageIndex')
	})

	test('setDrawingText is exposed with drawing-object selector guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setDrawingText')
		expect(schema?.schema.required).toEqual(['op', 'sheet', 'text'])
		expect(schema?.schema.properties.drawingObjectIndex?.description).toContain('drawing object')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setDrawingText',
			sheet: 'Sheet1',
			drawingPartPath: 'xl/drawings/drawing1.xml',
			id: 2,
			text: 'Updated callout',
		})
		expect(schema?.recoveryActions.join('\n')).toContain('visualInventory')
	})

	test('setThreadedComment is exposed with threaded comment selector guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setThreadedComment')
		expect(schema?.schema.required).toEqual(['op', 'sheet', 'text'])
		expect(schema?.schema.properties.threadedCommentId?.description).toContain('Threaded comment')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setThreadedComment',
			sheet: 'Sheet1',
			threadedCommentId: '{thread-id}',
			text: 'Updated review note',
		})
		expect(schema?.recoveryActions.join('\n')).toContain('threadedComments')
	})

	test('setChartSeriesSource is exposed with chart selector guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setChartSeriesSource')
		expect(schema?.schema.required).toEqual(['op', 'seriesIndex'])
		expect(schema?.schema.properties.partPath?.description).toContain('part path')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setChartSeriesSource',
			partPath: 'xl/charts/chart1.xml',
			seriesIndex: 0,
		})
		expect(schema?.recoveryActions.join('\n')).toContain('visualInventory')
	})

	test('setPivotCache is exposed with refresh guidance for analytics edits', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setPivotCache')
		expect(schema?.schema.required).toEqual(['op'])
		expect(schema?.schema.properties.pivotTable?.description).toContain('Pivot table')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setPivotCache',
			pivotTable: 'PivotTable1',
			refreshOnLoad: true,
		})
		expect(schema?.recoveryActions.join('\n')).toContain('invalid=true')
	})

	test('setAutoFilter is exposed with criteria-preserving edit guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setAutoFilter')
		expect(schema?.schema.required).toEqual(['op', 'sheet', 'range'])
		expect(schema?.schema.properties.values?.description).toContain('Filter value-list')
		expect(schema?.schema.properties.sortBy?.description).toContain('sort condition')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setAutoFilter',
			sheet: 'Sheet1',
			range: 'A1:D20',
			column: 0,
			values: ['North'],
		})

		const parsed = parseOperations([
			{
				op: 'setAutoFilter',
				sheet: 'Sheet1',
				range: 'A1:D20',
				column: 0,
				values: ['North'],
				sortBy: 'A2:A20',
				descending: true,
			},
		])
		expect(parsed.ok).toBe(true)
	})

	test('table topology operations expose ownership preconditions', () => {
		const create = getOperationsSchema().find((entry) => entry.op === 'createTable')
		const append = getOperationsSchema().find((entry) => entry.op === 'appendRows')
		const rename = getOperationsSchema().find((entry) => entry.op === 'renameTable')
		const resize = getOperationsSchema().find((entry) => entry.op === 'resizeTable')

		expect(create?.description).toContain('non-overlapping')
		expect(create?.schema.properties.name?.description).toContain('workbook-unique')
		expect(create?.recoveryActions.join('\n')).toContain('overlapping table ranges')
		expect(append?.description).toContain('shifting another table')
		expect(append?.recoveryActions.join('\n')).toContain('totals-row appends')
		expect(rename?.description).toContain('workbook-unique')
		expect(rename?.recoveryActions.join('\n')).toContain('case-insensitively')
		expect(rename?.recoveryActions.join('\n')).toContain('R1C1-style')
		expect(resize?.description).toContain('dropping referenced fields')
		expect(resize?.schema.properties.table?.description).toContain('Workbook-unique table name')
	})

	test('setTableStyle is exposed with style-name and flag validation', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setTableStyle')
		expect(schema?.schema.required).toEqual(['op', 'table'])
		expect(schema?.schema.properties.styleName?.type).toEqual(['string', 'null'])
		expect(schema?.schema.properties.showRowStripes?.type).toBe('boolean')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setTableStyle',
			table: 'Sales',
			styleName: 'TableStyleMedium2',
			showRowStripes: true,
		})

		const parsed = parseOperations([
			{
				op: 'setTableStyle',
				table: 'Sales',
				styleName: null,
				showFirstColumn: false,
				showLastColumn: true,
				showRowStripes: true,
				showColumnStripes: false,
			},
		])
		expect(parsed.ok).toBe(true)

		const invalid = parseOperations([
			{
				op: 'setTableStyle',
				table: 'Sales',
				showRowStripes: 'yes',
			},
		])
		expect(invalid.ok).toBe(false)
		if (!invalid.ok) expect(invalid.issues[0]).toContain('showRowStripes must be a boolean')
	})

	test('setTableColumn is exposed with column rename guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setTableColumn')
		expect(schema?.schema.required).toEqual(['op', 'table', 'column'])
		expect(schema?.schema.properties.newName?.type).toBe('string')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setTableColumn',
			table: 'Sales',
			column: 'Total',
			newName: 'Line Total',
		})
		expect(schema?.recoveryActions.join('\n')).toContain('rewrite structured references')

		const parsed = parseOperations([
			{
				op: 'setTableColumn',
				table: 'Sales',
				column: 'Qty',
				newName: 'Units',
				formula: null,
				totalsRowFormula: null,
			},
		])
		expect(parsed.ok).toBe(true)
	})

	test('setPivotFieldItem is exposed with filter item guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setPivotFieldItem')
		expect(schema?.schema.required).toEqual(['op', 'fieldIndex', 'itemIndex'])
		expect(schema?.schema.properties.selectedPageItem?.description).toContain('page-field')
		expect(schema?.schema.properties.hidden?.type).toEqual(['boolean', 'null'])
		expect(schema?.schema.properties.showDetails?.type).toEqual(['boolean', 'null'])
		expect(schema?.schema.properties.manualFilter?.type).toEqual(['boolean', 'null'])
		expect(schema?.schema.properties.selectedPageItem?.type).toEqual(['integer', 'null'])
		expect(schema?.examples[0]).toMatchObject({
			op: 'setPivotFieldItem',
			pivotTable: 'PivotTable1',
			fieldIndex: 0,
			itemIndex: 2,
			hidden: true,
		})
		expect(schema?.recoveryActions.join('\n')).toContain('fieldIndex')

		const parsed = parseOperations([
			{
				op: 'setPivotFieldItem',
				pivotTable: 'PivotTable1',
				fieldIndex: 0,
				itemIndex: 2,
				hidden: null,
				selectedPageItem: null,
			},
		])
		expect(parsed.ok).toBe(true)
	})

	test('setSlicerCacheItem is exposed with slicer refresh guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setSlicerCacheItem')
		expect(schema?.schema.required).toEqual(['op', 'item'])
		expect(schema?.schema.properties.slicerCache?.description).toContain('Slicer cache')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setSlicerCacheItem',
			slicerCache: 'Slicer_State',
			item: 0,
			selected: true,
		})
		expect(schema?.recoveryActions.join('\n')).toContain('pivot output')

		const parsed = parseOperations([
			{ op: 'setSlicerCacheItem', slicerCache: 'Slicer_State', item: 0, selected: null },
		])
		expect(parsed.ok).toBe(true)
	})

	test('setTimelineRange is exposed with timeline refresh guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setTimelineRange')
		expect(schema?.schema.required).toEqual(['op', 'startDate', 'endDate'])
		expect(schema?.schema.properties.timelineCache?.description).toContain('Timeline cache')
		expect(schema?.schema.properties.startDate?.description).toContain('start date-time')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setTimelineRange',
			timelineCache: 'Timeline_Order_Date',
			startDate: '2024-01-01T00:00:00',
			endDate: '2024-03-31T00:00:00',
		})
		expect(schema?.recoveryActions.join('\n')).toContain('pivot output')

		const parsed = parseOperations([
			{
				op: 'setTimelineRange',
				timelineCache: 'Timeline_Order_Date',
				startDate: '2024-01-01T00:00:00',
				endDate: '2024-03-31T00:00:00',
			},
		])
		expect(parsed.ok).toBe(true)
	})

	test('setSparklineGroup is exposed with sparkline range guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setSparklineGroup')
		expect(schema?.schema.required).toEqual(['op', 'sheet', 'groupIndex'])
		expect(schema?.schema.properties.groupIndex?.description).toContain('sparkline group')
		expect(schema?.schema.properties.locationRange?.description).toContain('sqref')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setSparklineGroup',
			sheet: 'Data',
			groupIndex: 0,
			range: 'Data!C2:C4',
			locationRange: 'E2:E4',
			type: 'column',
			markers: false,
		})
		expect(schema?.recoveryActions.join('\n')).toContain('sparklineGroups')

		const parsed = parseOperations([
			{
				op: 'setSparklineGroup',
				sheet: 'Data',
				groupIndex: 0,
				range: 'Data!C2:C4',
				locationRange: 'E2:E4',
				type: 'column',
				markers: false,
			},
		])
		expect(parsed.ok).toBe(true)
	})

	test('setAdvancedFilter is exposed with custom sheet view guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setAdvancedFilter')
		expect(schema?.schema.required).toEqual(['op', 'sheet', 'filterIndex'])
		expect(schema?.schema.properties.filterIndex?.description).toContain('advanced filter')
		expect(schema?.schema.properties.values?.description).toContain('Filter value-list')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setAdvancedFilter',
			sheet: 'Data',
			filterIndex: 0,
			range: 'A1:D20',
			column: 0,
			values: ['East', 'North'],
			sortRef: 'A2:D20',
			sortBy: 'B2:B20',
			descending: false,
		})
		expect(schema?.recoveryActions.join('\n')).toContain('advancedFilters')

		const parsed = parseOperations([
			{
				op: 'setAdvancedFilter',
				sheet: 'Data',
				filterIndex: 0,
				column: 0,
				values: ['East', 'North'],
				descending: false,
			},
		])
		expect(parsed.ok).toBe(true)
	})

	test('setConnectionRefresh is exposed with connection refresh guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setConnectionRefresh')
		expect(schema?.schema.required).toEqual(['op'])
		expect(schema?.schema.properties.connectionId?.description).toContain('connection id')
		expect(schema?.schema.properties.refreshedVersion?.description).toContain('refresh engine')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setConnectionRefresh',
			partPath: 'xl/queryTables/queryTable1.xml',
			connectionId: 1,
			refreshOnLoad: true,
			saveData: false,
		})
		expect(schema?.recoveryActions.join('\n')).toContain('refreshMetadata')

		const parsed = parseOperations([
			{
				op: 'setConnectionRefresh',
				partPath: 'xl/queryTables/queryTable1.xml',
				connectionId: 1,
				refreshOnLoad: true,
				saveData: false,
				refreshedVersion: 9,
			},
		])
		expect(parsed.ok).toBe(true)
	})

	test('rewriteExternalLink is exposed with external reference selector guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'rewriteExternalLink')
		expect(schema?.schema.required).toEqual(['op', 'newTarget'])
		expect(schema?.schema.properties.linkRelId?.description).toContain('external link part')
		expect(schema?.examples[0]).toMatchObject({
			op: 'rewriteExternalLink',
			partPath: 'xl/externalLinks/externalLink1.xml',
			newTarget: '../sources/reforecast.xlsx',
		})
		expect(schema?.recoveryActions.join('\n')).toContain('external-refs')
	})

	test('setWorkbookProperties is exposed with merge and clear guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setWorkbookProperties')
		expect(schema?.schema.required).toEqual(['op', 'properties'])
		expect(schema?.schema.properties.properties?.description).toContain('date1904')
		expect(schema?.schema.properties.mode?.enum).toContain('merge')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setWorkbookProperties',
			properties: { codeName: 'Model', filterPrivacy: true, date1904: false },
			mode: 'merge',
		})
		expect(schema?.recoveryActions.join('\n')).toContain('null property values')
	})

	test('setDocumentProperties is exposed with docProps merge guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setDocumentProperties')
		expect(schema?.schema.required).toEqual(['op', 'properties'])
		expect(schema?.schema.properties.mode?.enum).toContain('merge')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setDocumentProperties',
			properties: {
				core: { title: 'Forecast Pack', creator: 'Finance Ops' },
				app: { HeadingPairs: ['Worksheets', 1], TitlesOfParts: ['Sheet1'] },
				custom: [{ name: 'Reviewed', value: true }],
			},
			mode: 'merge',
		})
		expect(schema?.recoveryActions.join('\n')).toContain('docProps edits')
	})

	test('setTheme is exposed with theme color guidance', () => {
		const schema = getOperationsSchema().find((entry) => entry.op === 'setTheme')
		expect(schema?.schema.required).toEqual(['op'])
		expect(schema?.schema.properties.themeColors?.description).toContain('accent1-6')
		expect(schema?.examples[0]).toMatchObject({
			op: 'setTheme',
			themeName: 'Brand Theme',
			colorSchemeName: 'Brand Colors',
			themeColors: [{ slot: 'accent1', rgb: '0F6CBD' }],
		})
		expect(schema?.recoveryActions.join('\n')).toContain('themeSummary')
	})

	test('setWorkbookView and setCalcSettings expose workbook metadata guidance', () => {
		const view = getOperationsSchema().find((entry) => entry.op === 'setWorkbookView')
		expect(view?.schema.required).toEqual(['op', 'view'])
		expect(view?.schema.properties.view?.description).toContain('activeTab')
		expect(view?.examples[0]).toMatchObject({
			op: 'setWorkbookView',
			index: 0,
			view: { activeTab: 0, firstSheet: 0 },
			mode: 'merge',
		})
		expect(view?.recoveryActions.join('\n')).toContain('primary workbook view')

		const calc = getOperationsSchema().find((entry) => entry.op === 'setCalcSettings')
		expect(calc?.schema.required).toEqual(['op', 'settings'])
		expect(calc?.schema.properties.settings?.description).toContain('dateSystem')
		expect(calc?.examples[0]).toMatchObject({
			op: 'setCalcSettings',
			settings: { calcMode: 'manual', fullCalcOnLoad: true },
		})
		expect(calc?.recoveryActions.join('\n')).toContain('dateSystem carefully')
	})
})
