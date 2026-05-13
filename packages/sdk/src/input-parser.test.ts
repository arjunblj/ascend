import { describe, expect, test } from 'bun:test'
import { parseCellInput, parseCellInputOperation } from './index.ts'

describe('parseCellInput', () => {
	test('normalizes blank, scalar, escaped text, and formula input', () => {
		expect(parseCellInput('')).toEqual({ kind: 'blank', value: null, reason: 'blank' })
		expect(parseCellInput("'=not a formula")).toEqual({
			kind: 'value',
			value: '=not a formula',
			valueKind: 'string',
			reason: 'escaped-text',
		})
		expect(parseCellInput('42')).toEqual({
			kind: 'value',
			value: 42,
			valueKind: 'number',
			reason: 'number',
		})
		expect(parseCellInput('12.5%')).toEqual({
			kind: 'value',
			value: 0.125,
			valueKind: 'number',
			reason: 'percent',
		})
		expect(parseCellInput('TRUE')).toEqual({
			kind: 'value',
			value: true,
			valueKind: 'boolean',
			reason: 'boolean',
		})
		expect(parseCellInput('=SUM(A1:A2)')).toEqual({
			kind: 'formula',
			formula: 'SUM(A1:A2)',
			parseOk: true,
			reason: 'formula',
		})
	})

	test('surfaces formula parse errors without coercing the input to text', () => {
		const parsed = parseCellInput('=SUM(')

		expect(parsed.kind).toBe('formula')
		expect(parsed.parseOk).toBe(false)
		expect(parsed.parseError).toBeString()
	})

	test('parses ISO dates only when date parsing is requested', () => {
		expect(parseCellInput('2024-01-02')).toEqual({
			kind: 'value',
			value: '2024-01-02',
			valueKind: 'string',
			reason: 'text',
		})

		const parsed = parseCellInput('2024-01-02', { parseDates: true })
		expect(parsed.kind).toBe('value')
		expect(parsed.valueKind).toBe('date')
		expect(parsed.value).toBeInstanceOf(Date)
	})

	test('compiles parsed input to canonical operations', () => {
		expect(parseCellInputOperation('Sheet1', 'A1', '=B1*2')).toEqual({
			op: 'setFormula',
			sheet: 'Sheet1',
			ref: 'A1',
			formula: 'B1*2',
		})
		expect(parseCellInputOperation('Sheet1', 'A1', '17')).toEqual({
			op: 'setCells',
			sheet: 'Sheet1',
			updates: [{ ref: 'A1', value: 17 }],
		})
	})
})
