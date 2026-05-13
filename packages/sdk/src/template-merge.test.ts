import { describe, expect, test } from 'bun:test'
import { AscendWorkbook } from './index.ts'

describe('template merge', () => {
	test('compiles string and formula placeholders into replayable operations', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: '{{amount}}' },
					{ ref: 'A2', value: 'Invoice for {{client}}' },
					{ ref: 'A3', value: 'Missing {{unknown}}' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1+{{tax}}' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: '{{label}}&A2' },
		])

		const merge = wb.templateMerge({
			amount: 125,
			client: 'Acme',
			tax: 10,
			label: 'Total: ',
		})

		expect(merge.replayable).toBe(false)
		expect(merge.sheetCount).toBe(1)
		expect(merge.cellCount).toBe(5)
		expect(merge.formulaCount).toBe(2)
		expect(merge.replacementCount).toBe(4)
		expect(merge.unresolved).toEqual([
			{
				sheet: 'Sheet1',
				ref: 'A3',
				source: 'value',
				placeholder: '{{unknown}}',
				key: 'unknown',
			},
		])
		expect(merge.unsupported).toEqual([])
		expect(merge.ops).toEqual([
			{
				op: 'setCells',
				sheet: 'Sheet1',
				updates: [
					{ ref: 'A1', value: 125 },
					{ ref: 'A2', value: 'Invoice for Acme' },
				],
			},
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: 'A1+10' },
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B2', formula: '"Total: "&A2' },
		])

		const replayed = AscendWorkbook.create()
		expect(replayed.batch(merge.ops).errors).toEqual([])
		expect(replayed.recalc().errors).toEqual([])
		expect(replayed.get('Sheet1!A1')).toEqual({ kind: 'number', value: 125 })
		expect(replayed.get('Sheet1!A2')).toEqual({ kind: 'string', value: 'Invoice for Acme' })
		expect(replayed.get('Sheet1!B1')).toEqual({ kind: 'number', value: 135 })
		expect(replayed.sheet('Sheet1')?.cell('B2')?.formula).toBe('"Total: "&A2')
	})

	test('reports unsupported rich-text placeholder contexts without lossy conversion', () => {
		const wb = AscendWorkbook.create()
		wb.apply([
			{
				op: 'setRichText',
				sheet: 'Sheet1',
				ref: 'C1',
				runs: [
					{ text: 'Prepared for ', bold: true },
					{ text: '{{client}}', italic: true },
				],
			},
		])

		const merge = wb.templateMerge({ client: 'Acme' })

		expect(merge.replayable).toBe(false)
		expect(merge.ops).toEqual([])
		expect(merge.unresolved).toEqual([])
		expect(merge.unsupported).toEqual([
			{
				sheet: 'Sheet1',
				ref: 'C1',
				source: 'value',
				valueKind: 'richText',
				reason: 'Rich text placeholders require run-preserving merge support before replay.',
			},
		])
	})
})
