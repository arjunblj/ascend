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
})
