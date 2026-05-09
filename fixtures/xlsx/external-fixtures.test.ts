import { describe, expect, it } from 'bun:test'
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'

const EXTERNAL_DIR = join(dirname(fileURLToPath(import.meta.url)), 'external')

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

function getExternalFixtures(): string[] {
	if (!existsSync(EXTERNAL_DIR)) return []
	return readdirSync(EXTERNAL_DIR)
		.filter((name) => name.endsWith('.xlsx') || name.endsWith('.xlsm'))
		.sort((a, b) => a.localeCompare(b))
}

function loadExternalFixture(name: string): Uint8Array {
	return new Uint8Array(readFileSync(join(EXTERNAL_DIR, name)))
}

function countCells(workbook: {
	sheets: readonly { cells: { iterate: () => Iterable<unknown> } }[]
}): number {
	let n = 0
	for (const sheet of workbook.sheets) {
		for (const _ of sheet.cells.iterate()) n++
	}
	return n
}

function summarizeWorkbook(workbook: {
	sheets: readonly {
		merges: readonly unknown[]
		cells: { iterate: () => Iterable<unknown> }
	}[]
	definedNames: { list: () => readonly unknown[] }
	themeMetadata: { colorCount: number }
}): {
	sheetCount: number
	cellCount: number
	mergeCount: number
	definedNameCount: number
	themeColorCount: number
} {
	return {
		sheetCount: workbook.sheets.length,
		cellCount: countCells(workbook),
		mergeCount: workbook.sheets.reduce((sum, sheet) => sum + sheet.merges.length, 0),
		definedNameCount: workbook.definedNames.list().length,
		themeColorCount: workbook.themeMetadata.colorCount,
	}
}

if (getExternalFixtures().length > 0) {
	describe('External XLSX fixtures', () => {
		for (const fixture of getExternalFixtures()) {
			it(`reads ${fixture} without errors`, () => {
				const result = readXlsx(loadExternalFixture(fixture))
				expectOk(result)
				expect(result.value.workbook.sheets.length).toBeGreaterThan(0)
			})

			it(`verifies basic properties of ${fixture}`, () => {
				const result = readXlsx(loadExternalFixture(fixture))
				expectOk(result)
				const wb = result.value.workbook
				expect(wb.sheets.length).toBeGreaterThan(0)
				const cellCount = countCells(wb)
				expect(cellCount).toBeGreaterThan(0)
			})

			it(`round-trips ${fixture}`, () => {
				const initial = readXlsx(loadExternalFixture(fixture))
				expectOk(initial)
				const before = summarizeWorkbook(initial.value.workbook)
				const written = writeXlsx(initial.value.workbook, initial.value.capsules)
				expectOk(written)
				const reopened = readXlsx(written.value)
				expectOk(reopened)
				expect(summarizeWorkbook(reopened.value.workbook)).toEqual(before)
			})
		}

		it('keeps UK government spend workbook structurally rich', () => {
			const result = readXlsx(loadExternalFixture('uk-gov-spend-nice-2026-02.xlsx'))
			expectOk(result)
			expect(summarizeWorkbook(result.value.workbook)).toMatchObject({
				sheetCount: 2,
				mergeCount: 2,
				definedNameCount: 13,
				themeColorCount: 12,
			})
		})

		it('keeps Census workbook large and merge-heavy', () => {
			const result = readXlsx(loadExternalFixture('us-census-construction-2025-10.xlsx'))
			expectOk(result)
			const summary = summarizeWorkbook(result.value.workbook)
			expect(summary.sheetCount).toBe(3)
			expect(summary.cellCount).toBeGreaterThan(1000)
			expect(summary.mergeCount).toBeGreaterThanOrEqual(10)
			expect(summary.definedNameCount).toBeGreaterThanOrEqual(2)
		})

		it('keeps SEC workbook multi-sheet and dense', () => {
			const result = readXlsx(loadExternalFixture('sec-mmf-statistics-2022-02.xlsx'))
			expectOk(result)
			const summary = summarizeWorkbook(result.value.workbook)
			expect(summary.sheetCount).toBeGreaterThanOrEqual(10)
			expect(summary.cellCount).toBeGreaterThan(5000)
			expect(summary.themeColorCount).toBe(12)
		})
	})
} else {
	describe('External XLSX fixtures', () => {
		it('skips when fixtures not downloaded (run fixtures/xlsx/download-fixtures.sh)', () => {
			expect(getExternalFixtures()).toHaveLength(0)
		})
	})
}
