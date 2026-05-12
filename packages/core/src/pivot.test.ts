import { describe, expect, test } from 'bun:test'
import { pivotDataFieldCaptionsMatch } from './pivot.ts'

describe('pivot data field captions', () => {
	test('matches Excel GETPIVOTDATA root names against aggregate captions', () => {
		expect(pivotDataFieldCaptionsMatch('Sales', 'Sum of Sales')).toBe(true)
		expect(pivotDataFieldCaptionsMatch('Sales', 'Count of Sales')).toBe(true)
		expect(pivotDataFieldCaptionsMatch('Sales', 'Average of Sales')).toBe(true)
		expect(pivotDataFieldCaptionsMatch('Sum of Sales', 'Sales')).toBe(true)
		expect(pivotDataFieldCaptionsMatch('Sales', 'Profit')).toBe(false)
	})
})
