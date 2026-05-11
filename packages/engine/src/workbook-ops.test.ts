import { describe, expect, test } from 'bun:test'
import { createWorkbook } from '@ascend/core'
import { applyOperation } from './operations.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

function expectErr<T, E>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: false; error: E } {
	expect(result.ok).toBe(false)
	if (result.ok) throw new Error('Expected operation to fail')
}

describe('workbook metadata operations', () => {
	test('setWorkbookProperties merges and clears workbookPr metadata', () => {
		const wb = createWorkbook()
		wb.workbookProperties = {
			codeName: 'OldModel',
			defaultThemeVersion: 166925,
			filterPrivacy: false,
		}

		const result = applyOperation(wb, {
			op: 'setWorkbookProperties',
			properties: {
				codeName: 'Model',
				defaultThemeVersion: null,
				filterPrivacy: true,
				date1904: true,
			},
		})
		expectOk(result)

		expect(wb.workbookProperties).toEqual({
			codeName: 'Model',
			filterPrivacy: true,
			date1904: true,
		})
		expect(wb.calcSettings.dateSystem).toBe('1904')
		expect(result.value.recalcRequired).toBe(true)
	})

	test('setWorkbookProperties replaces workbookPr metadata', () => {
		const wb = createWorkbook()
		wb.workbookProperties = { codeName: 'OldModel', filterPrivacy: true, date1904: true }
		wb.calcSettings = { ...wb.calcSettings, dateSystem: '1904' }

		const result = applyOperation(wb, {
			op: 'setWorkbookProperties',
			mode: 'replace',
			properties: { codeName: 'NewModel' },
		})
		expectOk(result)

		expect(wb.workbookProperties).toEqual({ codeName: 'NewModel' })
		expect(wb.calcSettings.dateSystem).toBe('1900')
		expect(result.value.recalcRequired).toBe(true)
	})

	test('setWorkbookProperties validates property values', () => {
		const wb = createWorkbook()
		const result = applyOperation(wb, {
			op: 'setWorkbookProperties',
			properties: { defaultThemeVersion: -1 },
		})

		expectErr(result)
		expect(result.error.message).toContain('defaultThemeVersion')
	})

	test('setDocumentProperties merges core, app, and custom docProps metadata', () => {
		const wb = createWorkbook()
		wb.documentProperties = {
			core: { title: 'Old', creator: 'Analyst', subject: 'Forecast' },
			app: { Application: 'Excel', Company: 'OldCo' },
			custom: [{ name: 'Reviewed', value: false, type: 'bool', pid: 2 }],
		}

		const result = applyOperation(wb, {
			op: 'setDocumentProperties',
			properties: {
				core: { title: 'Board Pack', subject: null },
				app: { Company: 'Ascend', HeadingPairs: ['Worksheets', 1] },
				custom: [{ name: 'Reviewed', value: true, type: 'bool' }],
			},
		})
		expectOk(result)

		expect(wb.documentProperties).toEqual({
			core: { title: 'Board Pack', creator: 'Analyst' },
			app: { Application: 'Excel', Company: 'Ascend', HeadingPairs: ['Worksheets', 1] },
			custom: [{ name: 'Reviewed', value: true, type: 'bool' }],
		})
		expect(result.value.recalcRequired).toBe(false)
	})

	test('setDocumentProperties replaces or clears docProps families', () => {
		const wb = createWorkbook()
		wb.documentProperties = {
			core: { title: 'Old' },
			app: { Application: 'Excel' },
			custom: [{ name: 'Reviewed', value: true }],
		}

		const result = applyOperation(wb, {
			op: 'setDocumentProperties',
			mode: 'replace',
			properties: { core: { creator: 'Finance' }, app: null, custom: null },
		})
		expectOk(result)

		expect(wb.documentProperties).toEqual({ core: { creator: 'Finance' } })
	})

	test('setWorkbookView merges, appends, and deletes view metadata', () => {
		const wb = createWorkbook()
		wb.workbookViews.push({ activeTab: 0, firstSheet: 0, visibility: 'visible' })

		const merged = applyOperation(wb, {
			op: 'setWorkbookView',
			index: 0,
			view: { activeTab: 2, visibility: null },
		})
		expectOk(merged)
		expect(wb.workbookViews[0]).toEqual({ activeTab: 2, firstSheet: 0 })

		const appended = applyOperation(wb, {
			op: 'setWorkbookView',
			index: 1,
			mode: 'replace',
			view: { activeTab: 1, tabRatio: 600 },
		})
		expectOk(appended)
		expect(wb.workbookViews[1]).toEqual({ activeTab: 1, tabRatio: 600 })

		const deleted = applyOperation(wb, { op: 'setWorkbookView', index: 0, view: null })
		expectOk(deleted)
		expect(wb.workbookViews).toEqual([{ activeTab: 1, tabRatio: 600 }])
	})

	test('setCalcSettings updates calculation metadata and date system coherently', () => {
		const wb = createWorkbook()
		const result = applyOperation(wb, {
			op: 'setCalcSettings',
			settings: {
				calcMode: 'manual',
				fullCalcOnLoad: true,
				calcCompleted: null,
				dateSystem: '1904',
				iterativeCalc: { enabled: true, maxIterations: 50, maxChange: 0.0001 },
			},
		})
		expectOk(result)

		expect(wb.calcSettings).toMatchObject({
			calcMode: 'manual',
			fullCalcOnLoad: true,
			dateSystem: '1904',
			iterativeCalc: { enabled: true, maxIterations: 50, maxChange: 0.0001 },
		})
		expect(wb.calcSettings.calcCompleted).toBeUndefined()
		expect(wb.workbookProperties.date1904).toBe(true)
		expect(result.value.recalcRequired).toBe(true)
	})

	test('setTheme updates names, fonts, and color slots without recalculation', () => {
		const wb = createWorkbook()
		wb.themeMetadata = {
			name: 'Office',
			colorSchemeName: 'Office',
			colorCount: 2,
			majorFontLatin: 'Aptos Display',
			minorFontLatin: 'Aptos',
		}
		wb.themeColors.push(
			{ slot: 'dk1', systemColor: 'windowText', lastColor: '000000' },
			{ slot: 'accent1', rgb: '4F81BD' },
		)

		const result = applyOperation(wb, {
			op: 'setTheme',
			themeName: 'Brand Theme',
			colorSchemeName: 'Brand Colors',
			majorFontLatin: 'Inter Display',
			minorFontLatin: 'Inter',
			themeColors: [
				{ slot: 'accent1', rgb: '0f6cbd' },
				{ slot: 'lt1', systemColor: 'window', lastColor: 'ffffff' },
			],
		})
		expectOk(result)

		expect(result.value.recalcRequired).toBe(false)
		expect(wb.themeMetadata).toEqual({
			name: 'Brand Theme',
			colorSchemeName: 'Brand Colors',
			colorCount: 3,
			majorFontLatin: 'Inter Display',
			minorFontLatin: 'Inter',
		})
		expect(wb.themeColors).toEqual([
			{ slot: 'dk1', systemColor: 'windowText', lastColor: '000000' },
			{ slot: 'accent1', rgb: '0F6CBD' },
			{ slot: 'lt1', systemColor: 'window', lastColor: 'FFFFFF' },
		])
	})

	test('setTheme validates color slots and values', () => {
		const wb = createWorkbook()

		const missing = applyOperation(wb, { op: 'setTheme' })
		expectErr(missing)
		expect(missing.error.message).toContain('requires theme metadata')

		const badSlot = applyOperation(wb, {
			op: 'setTheme',
			themeColors: [{ slot: 'brand', rgb: '123456' }],
		})
		expectErr(badSlot)
		expect(badSlot.error.message).toContain('Unsupported theme color slot')

		const badRgb = applyOperation(wb, {
			op: 'setTheme',
			themeColors: [{ slot: 'accent1', rgb: '#123456' }],
		})
		expectErr(badRgb)
		expect(badRgb.error.message).toContain('6 hex digits')
	})
})
