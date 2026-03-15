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
				const written = writeXlsx(initial.value.workbook, initial.value.capsules)
				expectOk(written)
				const reopened = readXlsx(written.value)
				expectOk(reopened)
				expect(reopened.value.workbook.sheets.length).toBe(initial.value.workbook.sheets.length)
			})
		}
	})
} else {
	describe('External XLSX fixtures', () => {
		it('skips when fixtures not downloaded (run fixtures/xlsx/download-fixtures.sh)', () => {
			expect(getExternalFixtures()).toHaveLength(0)
		})
	})
}
