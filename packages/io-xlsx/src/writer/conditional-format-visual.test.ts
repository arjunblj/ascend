import { describe, expect, test } from 'bun:test'
import { type StyleId, Workbook } from '@ascend/core'
import { numberValue } from '@ascend/schema'
import { readXlsx } from '../reader/index.ts'
import { writeXlsx } from './index.ts'

const S0 = 0 as StyleId

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	if (!result.ok) throw new Error(result.error.message)
	expect(result.ok).toBe(true)
}

describe('visual conditional formatting', () => {
	test('round-trips color scales, data bars, and icon sets', () => {
		const wb = new Workbook()
		const sheet = wb.addSheet('Data')
		for (let row = 0; row < 5; row++) {
			sheet.cells.set(row, 0, { value: numberValue(row + 1), formula: null, styleId: S0 })
		}
		sheet.conditionalFormats.push(
			{
				sqref: 'A1:A5',
				rules: [
					{
						type: 'colorScale',
						priority: 1,
						formulas: [],
						colorScale: {
							cfvo: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
							colors: [{ rgb: 'FFF8696B' }, { rgb: 'FFFFEB84' }, { rgb: 'FF63BE7B' }],
						},
					},
				],
			},
			{
				sqref: 'B1:B5',
				rules: [
					{
						type: 'dataBar',
						priority: 2,
						formulas: [],
						dataBar: {
							cfvo: [{ type: 'min' }, { type: 'max' }],
							color: { rgb: 'FF638EC6' },
							showValue: false,
							minLength: 10,
							maxLength: 90,
						},
					},
				],
			},
			{
				sqref: 'C1:C5',
				rules: [
					{
						type: 'iconSet',
						priority: 3,
						formulas: [],
						iconSet: {
							iconSet: '3TrafficLights1',
							cfvo: [
								{ type: 'percent', value: '0' },
								{ type: 'percent', value: '33', gte: false },
								{ type: 'percent', value: '67' },
							],
							showValue: true,
							percent: true,
							reverse: false,
						},
					},
				],
			},
		)

		const written = writeXlsx(wb)
		expectOk(written)
		const reopened = readXlsx(written.value)
		expectOk(reopened)

		expect(reopened.value.workbook.sheets[0]?.conditionalFormats).toEqual(sheet.conditionalFormats)
	})
})
