import { describe, expect, test } from 'bun:test'
import {
	booleanValue,
	EMPTY,
	errorValue,
	isEmpty,
	isError,
	numberValue,
	stringValue,
} from './values.ts'

describe('CellValue constructors', () => {
	test('EMPTY has kind empty', () => {
		expect(EMPTY.kind).toBe('empty')
		expect(isEmpty(EMPTY)).toBe(true)
	})

	test('numberValue creates number cell', () => {
		const v = numberValue(42)
		expect(v).toEqual({ kind: 'number', value: 42 })
		expect(isEmpty(v)).toBe(false)
	})

	test('stringValue creates string cell', () => {
		const v = stringValue('hello')
		expect(v).toEqual({ kind: 'string', value: 'hello' })
	})

	test('booleanValue creates boolean cell', () => {
		expect(booleanValue(true)).toEqual({ kind: 'boolean', value: true })
		expect(booleanValue(false)).toEqual({ kind: 'boolean', value: false })
	})

	test('errorValue creates error cell', () => {
		const v = errorValue('#DIV/0!')
		expect(v).toEqual({ kind: 'error', value: '#DIV/0!' })
		expect(isError(v)).toBe(true)
	})

	test('isError returns false for non-errors', () => {
		expect(isError(numberValue(1))).toBe(false)
		expect(isError(EMPTY)).toBe(false)
	})
})
