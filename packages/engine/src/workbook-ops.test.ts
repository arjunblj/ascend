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
})
