import { describe, expect, test } from 'bun:test'
import { MACHINE_FORMAT_VERSION, machineFailure, machineSuccess } from './machine.ts'

describe('machine envelopes', () => {
	test('machineSuccess wraps data with stable version', () => {
		expect(machineSuccess({ value: 1 })).toEqual({
			formatVersion: MACHINE_FORMAT_VERSION,
			ok: true,
			data: { value: 1 },
		})
	})

	test('machineFailure wraps string errors', () => {
		expect(machineFailure('bad')).toEqual({
			formatVersion: MACHINE_FORMAT_VERSION,
			ok: false,
			error: { message: 'bad' },
		})
	})

	test('machineFailure preserves structured error detail', () => {
		expect(
			machineFailure({
				code: 'SHEET_NOT_FOUND',
				message: 'Sheet not found',
				retryable: false,
				retryStrategy: 'modified',
				refs: ['Sheet1!A1'],
				details: { sheet: 'Sheet1' },
				suggestedFix: 'Use a valid sheet name',
			}),
		).toEqual({
			formatVersion: MACHINE_FORMAT_VERSION,
			ok: false,
			error: {
				code: 'SHEET_NOT_FOUND',
				message: 'Sheet not found',
				retryable: false,
				retryStrategy: 'modified',
				refs: ['Sheet1!A1'],
				details: { sheet: 'Sheet1' },
				suggestedFix: 'Use a valid sheet name',
			},
		})
	})
})
