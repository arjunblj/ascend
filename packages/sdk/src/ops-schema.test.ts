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
})
