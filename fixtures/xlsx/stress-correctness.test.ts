import { describe, expect, test } from 'bun:test'
import { createWorkbook, type StyleId } from '../../packages/core/src/index.ts'
import { defaultCalcContext, recalculate } from '../../packages/engine/src/index.ts'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { EMPTY, numberValue, stringValue } from '../../packages/schema/src/index.ts'

const sid = 0 as StyleId

describe('large-scale stress correctness', () => {
	test('100K-cell workbook roundtrips correctly', () => {
		const ROWS = 1000
		const COLS = 100
		const wb = createWorkbook()
		const sheet = wb.addSheet('Data')
		for (let r = 0; r < ROWS; r++) {
			for (let c = 0; c < COLS; c++) {
				sheet.cells.set(r, c, {
					value: numberValue(r * COLS + c),
					formula: null,
					styleId: sid,
				})
			}
		}
		const result = writeXlsx(wb)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const loaded = readXlsx(result.value)
		expect(loaded.ok).toBe(true)
		if (!loaded.ok) return

		const loadedSheet = loaded.value.workbook.sheets[0]
		expect(loadedSheet).toBeDefined()
		if (!loadedSheet) return

		let mismatches = 0
		for (let r = 0; r < ROWS; r++) {
			for (let c = 0; c < COLS; c++) {
				const expected = r * COLS + c
				const actual = loadedSheet.cells.readNumber(r, c)
				if (actual !== expected) mismatches++
			}
		}
		expect(mismatches).toBe(0)
	})

	test('10K formulas recalculate correctly', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Sheet1')
		for (let r = 0; r < 100; r++) {
			for (let c = 0; c < 10; c++) {
				sheet.cells.set(r, c, {
					value: numberValue(r * 10 + c + 1),
					formula: null,
					styleId: sid,
				})
			}
		}
		for (let r = 100; r < 200; r++) {
			for (let c = 0; c < 10; c++) {
				const srcCol = String.fromCharCode(65 + c)
				const srcRow = r - 100 + 1
				sheet.cells.set(r, c, {
					value: EMPTY,
					formula: `${srcCol}${String(srcRow)}*2`,
					styleId: sid,
				})
			}
		}
		recalculate(wb, defaultCalcContext())

		let errors = 0
		for (let r = 100; r < 200; r++) {
			for (let c = 0; c < 10; c++) {
				const expected = ((r - 100) * 10 + c + 1) * 2
				const actual = sheet.cells.readNumber(r, c)
				if (actual !== expected) errors++
			}
		}
		expect(errors).toBe(0)
	})

	test('string data roundtrips through XLSX', () => {
		const wb = createWorkbook()
		const sheet = wb.addSheet('Strings')
		const testStrings = ['hello', '日本語', 'émoji 🎉', '', ' leading space', 'trailing ']
		for (let i = 0; i < testStrings.length; i++) {
			sheet.cells.set(i, 0, {
				value: stringValue(testStrings[i] ?? ''),
				formula: null,
				styleId: sid,
			})
		}
		const result = writeXlsx(wb)
		expect(result.ok).toBe(true)
		if (!result.ok) return

		const loaded = readXlsx(result.value)
		expect(loaded.ok).toBe(true)
		if (!loaded.ok) return

		const ls = loaded.value.workbook.sheets[0]
		expect(ls).toBeDefined()
		if (!ls) return

		for (let i = 0; i < testStrings.length; i++) {
			const expected = testStrings[i] ?? ''
			const actual = ls.cells.readString(i, 0)
			if (expected === '') {
				expect(actual === null || actual === '').toBe(true)
			} else {
				expect(actual).toBe(expected)
			}
		}
	})
})
