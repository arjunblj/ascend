import { describe, expect, test } from 'bun:test'
import { readCsv } from './reader.ts'
import { writeCsv } from './writer.ts'

function firstSheet(result: ReturnType<typeof readCsv>) {
	if (!result.ok) throw new Error(result.error.message)
	const sheet = result.value.sheets[0]
	if (!sheet) throw new Error('No sheets in workbook')
	return sheet
}

function cellValue(result: ReturnType<typeof readCsv>, row: number, col: number) {
	return firstSheet(result).cells.get(row, col)?.value
}

describe('readCsv', () => {
	test('parses simple CSV with numbers and strings', () => {
		const csv = 'name,age\nAlice,30\nBob,25'
		const result = readCsv(csv)
		expect(result.ok).toBe(true)

		expect(cellValue(result, 0, 0)).toEqual({ kind: 'string', value: 'name' })
		expect(cellValue(result, 0, 1)).toEqual({ kind: 'string', value: 'age' })
		expect(cellValue(result, 1, 0)).toEqual({ kind: 'string', value: 'Alice' })
		expect(cellValue(result, 1, 1)).toEqual({ kind: 'number', value: 30 })
		expect(cellValue(result, 2, 0)).toEqual({ kind: 'string', value: 'Bob' })
		expect(cellValue(result, 2, 1)).toEqual({ kind: 'number', value: 25 })
	})

	test('parses quoted fields containing commas', () => {
		const csv = 'name,address\nAlice,"123 Main St, Apt 4"'
		const result = readCsv(csv)
		expect(result.ok).toBe(true)

		expect(cellValue(result, 1, 1)).toEqual({
			kind: 'string',
			value: '123 Main St, Apt 4',
		})
	})

	test('parses quoted fields with embedded newlines', () => {
		const csv = 'a,b\n"line1\nline2",val'
		const result = readCsv(csv)
		expect(result.ok).toBe(true)

		expect(cellValue(result, 1, 0)).toEqual({
			kind: 'string',
			value: 'line1\nline2',
		})
		expect(cellValue(result, 1, 1)).toEqual({ kind: 'string', value: 'val' })
	})

	test('parses TSV with tab delimiter', () => {
		const tsv = 'x\ty\n1\t2'
		const result = readCsv(tsv, { delimiter: '\t' })
		expect(result.ok).toBe(true)

		expect(cellValue(result, 0, 0)).toEqual({ kind: 'string', value: 'x' })
		expect(cellValue(result, 0, 1)).toEqual({ kind: 'string', value: 'y' })
		expect(cellValue(result, 1, 0)).toEqual({ kind: 'number', value: 1 })
		expect(cellValue(result, 1, 1)).toEqual({ kind: 'number', value: 2 })
	})

	test('handles BOM in input', () => {
		const csv = '\uFEFFa,b\n1,2'
		const result = readCsv(csv)
		expect(result.ok).toBe(true)

		expect(cellValue(result, 0, 0)).toEqual({ kind: 'string', value: 'a' })
	})

	test('respects hasHeader option (data still goes into cells)', () => {
		const csv = 'h1,h2\nv1,v2'
		const result = readCsv(csv, { hasHeader: true })
		expect(result.ok).toBe(true)

		expect(cellValue(result, 0, 0)).toEqual({ kind: 'string', value: 'h1' })
		expect(cellValue(result, 1, 0)).toEqual({ kind: 'string', value: 'v1' })
	})

	test('detects numbers and booleans', () => {
		const csv = 'a,b,c,d\n42,3.14,true,false'
		const result = readCsv(csv)
		expect(result.ok).toBe(true)

		expect(cellValue(result, 1, 0)).toEqual({ kind: 'number', value: 42 })
		expect(cellValue(result, 1, 1)).toEqual({ kind: 'number', value: 3.14 })
		expect(cellValue(result, 1, 2)).toEqual({ kind: 'boolean', value: true })
		expect(cellValue(result, 1, 3)).toEqual({ kind: 'boolean', value: false })
	})

	test('handles empty cells', () => {
		const csv = 'a,,c\n,2,'
		const result = readCsv(csv)
		expect(result.ok).toBe(true)

		expect(cellValue(result, 0, 1)).toBeUndefined()
		expect(cellValue(result, 1, 0)).toBeUndefined()
		expect(cellValue(result, 1, 2)).toBeUndefined()
		expect(cellValue(result, 0, 0)).toEqual({ kind: 'string', value: 'a' })
		expect(cellValue(result, 1, 1)).toEqual({ kind: 'number', value: 2 })
	})

	test('custom delimiter and quote char', () => {
		const csv = "a|b\n'hello|world'|42"
		const result = readCsv(csv, { delimiter: '|', quote: "'", escape: "'" })
		expect(result.ok).toBe(true)

		expect(cellValue(result, 1, 0)).toEqual({
			kind: 'string',
			value: 'hello|world',
		})
		expect(cellValue(result, 1, 1)).toEqual({ kind: 'number', value: 42 })
	})
})

describe('writeCsv', () => {
	test('write -> read roundtrip preserves values', () => {
		const csv = 'name,age\nAlice,30\nBob,25'
		const readResult = readCsv(csv)
		expect(readResult.ok).toBe(true)
		if (!readResult.ok) return

		const writeResult = writeCsv(readResult.value)
		expect(writeResult.ok).toBe(true)
		if (!writeResult.ok) return

		const reread = readCsv(writeResult.value)
		expect(reread.ok).toBe(true)
		if (!reread.ok) return

		expect(cellValue(reread, 0, 0)).toEqual({ kind: 'string', value: 'name' })
		expect(cellValue(reread, 1, 0)).toEqual({ kind: 'string', value: 'Alice' })
		expect(cellValue(reread, 1, 1)).toEqual({ kind: 'number', value: 30 })
		expect(cellValue(reread, 2, 1)).toEqual({ kind: 'number', value: 25 })
	})

	test('quotes fields containing delimiter or newlines', () => {
		const readResult = readCsv('a\n"hello,world"')
		expect(readResult.ok).toBe(true)
		if (!readResult.ok) return

		const writeResult = writeCsv(readResult.value)
		expect(writeResult.ok).toBe(true)
		if (!writeResult.ok) return

		expect(writeResult.value).toContain('"hello,world"')
	})
})
