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

	test('copyFrom clones defined-name entries and extra attributes without aliasing', () => {
		const names = new DefinedNameCollection()
		names.set(
			'_xlnm._FilterDatabase',
			'Data!$A$1:$B$10',
			{ kind: 'sheet', sheetId: 'sheet-1' },
			{
				hidden: true,
				extraAttributes: [{ name: 'comment', value: 'Filter menu' }],
			},
		)

		const clone = names.clone()
		const clonedEntry = clone.resolve('_xlnm._FilterDatabase', 'sheet-1')
		expect(clonedEntry).toBeDefined()
		if (!clonedEntry) return

		;(clonedEntry.extraAttributes?.[0] as { value: string }).value = 'Changed'
		;(clonedEntry.scope as { sheetId: string }).sheetId = 'sheet-2'

		const originalEntry = names.resolve('_xlnm._FilterDatabase', 'sheet-1')
		expect(originalEntry?.extraAttributes?.[0]?.value).toBe('Filter menu')
		expect(originalEntry?.scope).toEqual({ kind: 'sheet', sheetId: 'sheet-1' })
		expect(clone.resolve('_xlnm._FilterDatabase', 'sheet-1')).toBe(clonedEntry)
	})

	test('add preserves duplicate names while keeping deterministic lookup', () => {
		const names = new DefinedNameCollection()
		names.add('_xlnm._FilterDatabase', 'Data!$A$1:$B$2', { kind: 'sheet', sheetId: 'sheet-1' })
		names.add('_xlnm._FilterDatabase', 'Data!$A$1:$B$3', { kind: 'sheet', sheetId: 'sheet-1' })

		expect(names.size).toBe(2)
		expect(names.list().map((entry) => entry.formula)).toEqual(['Data!$A$1:$B$2', 'Data!$A$1:$B$3'])
		expect(names.resolve('_xlnm._FilterDatabase', 'sheet-1')?.formula).toBe('Data!$A$1:$B$2')
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
