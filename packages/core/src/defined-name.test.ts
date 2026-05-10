import { describe, expect, test } from 'bun:test'
import { DefinedNameCollection } from './defined-name.ts'

describe('DefinedNameCollection', () => {
	test('resolve prefers local scope over workbook scope', () => {
		const names = new DefinedNameCollection()
		names.set('Rate', '0.1')
		names.set('Rate', '0.2', { kind: 'sheet', sheetId: 'sheet-1' })

		expect(names.resolve('Rate', 'sheet-1')?.formula).toBe('0.2')
		expect(names.resolve('Rate', 'sheet-2')?.formula).toBe('0.1')
	})

	test('delete removes entries from lookup indexes', () => {
		const names = new DefinedNameCollection()
		names.set('Budget', '10')
		expect(names.resolve('Budget')?.formula).toBe('10')
		expect(names.delete('Budget')).toBe(true)
		expect(names.resolve('Budget')).toBeUndefined()
	})

	test('copyFrom preserves lookup behavior without aliasing collection mutations', () => {
		const names = new DefinedNameCollection()
		names.set('Rate', '0.1')
		names.set('Rate', '0.2', { kind: 'sheet', sheetId: 'sheet-1' })

		const clone = new DefinedNameCollection()
		clone.copyFrom(names)
		expect(clone.resolve('Rate', 'sheet-1')?.formula).toBe('0.2')
		expect(clone.resolve('Rate', 'sheet-2')?.formula).toBe('0.1')

		clone.set('Rate', '0.3', { kind: 'sheet', sheetId: 'sheet-1' })
		expect(clone.resolve('Rate', 'sheet-1')?.formula).toBe('0.3')
		expect(names.resolve('Rate', 'sheet-1')?.formula).toBe('0.2')
	})

	test('preserves optional hidden metadata', () => {
		const names = new DefinedNameCollection()
		names.set(
			'_xlnm._FilterDatabase',
			'Data!$A$1:$B$10',
			{ kind: 'sheet', sheetId: 'sheet-1' },
			{ hidden: true },
		)

		const entry = names.resolve('_xlnm._FilterDatabase', 'sheet-1')
		expect(entry?.hidden).toBe(true)

		names.set('VisibleName', 'Data!$A$1')
		expect(names.resolve('VisibleName')?.hidden).toBeUndefined()
	})
})
