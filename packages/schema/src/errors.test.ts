import { describe, expect, test } from 'bun:test'
import { ascendError, err, ok } from './errors.ts'

describe('Result', () => {
	test('ok wraps a value', () => {
		const r = ok(42)
		expect(r.ok).toBe(true)
		if (r.ok) expect(r.value).toBe(42)
	})

	test('err wraps an error', () => {
		const e = ascendError('INVALID_REF', 'Bad reference')
		const r = err(e)
		expect(r.ok).toBe(false)
		if (!r.ok) {
			expect(r.error.code).toBe('INVALID_REF')
			expect(r.error.message).toBe('Bad reference')
			expect(r.error.retryable).toBe(false)
		}
	})
})

describe('ascendError', () => {
	test('creates error with defaults', () => {
		const e = ascendError('FORMULA_PARSE_ERROR', 'Unexpected token')
		expect(e.code).toBe('FORMULA_PARSE_ERROR')
		expect(e.retryable).toBe(false)
		expect(e.refs).toBeUndefined()
		expect(e.suggestedFix).toBeUndefined()
	})

	test('creates error with options', () => {
		const e = ascendError('SHEET_NOT_FOUND', 'Sheet "Q1" not found', {
			refs: ['Q1!A1'],
			suggestedFix: 'Check sheet name spelling',
			retryable: false,
		})
		expect(e.refs).toEqual(['Q1!A1'])
		expect(e.suggestedFix).toBe('Check sheet name spelling')
	})
})
