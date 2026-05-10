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
