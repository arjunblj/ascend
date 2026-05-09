import { describe, expect, test } from 'bun:test'
import { Workbook } from '@ascend/core'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	if (!result.ok) throw new Error(result.error.message)
	expect(result.ok).toBe(true)
}

describe('data validation fidelity', () => {
	test('round-trips prompts, errors, dropdown policy, and IME mode', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		sheet.dataValidations.push({
			sqref: 'A2:A20',
			type: 'list',
			formula1: '=Choices',
			allowBlank: false,
			showDropDown: true,
			showInputMessage: true,
			promptTitle: 'Pick one',
			prompt: 'Choose a value from the approved list.',
			showErrorMessage: true,
			errorTitle: 'Invalid value',
			error: 'Use one of the approved values.',
			errorStyle: 'warning',
			imeMode: 'hiragana',
		})

		const written = writeXlsx(wb)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)

		expect(reopened.value.workbook.sheets[0]?.dataValidations).toEqual(sheet.dataValidations)
	})
})
