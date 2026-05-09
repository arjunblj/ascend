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
