import { describe, expect, it } from 'bun:test'
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'

const fixtureDir = join(dirname(fileURLToPath(import.meta.url)), 'xlsxwriter')

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

function fixtureNames(): string[] {
	if (!existsSync(fixtureDir)) return []
	return readdirSync(fixtureDir)
		.filter((name) => name.endsWith('.xlsx'))
		.sort((a, b) => a.localeCompare(b))
}

function loadFixture(name: string): Uint8Array {
	return new Uint8Array(readFileSync(join(fixtureDir, name)))
}

if (fixtureNames().length > 0) {
	describe('XlsxWriter fixture corpus', () => {
		for (const name of fixtureNames()) {
			it(`reads ${name}`, () => {
				const result = readXlsx(loadFixture(name))
				expectOk(result)
				expect(result.value.workbook.sheets.length).toBeGreaterThan(0)
			})

			it(`round-trips ${name}`, () => {
				const initial = readXlsx(loadFixture(name))
				expectOk(initial)
				const written = writeXlsx(initial.value.workbook, initial.value.capsules)
				expectOk(written)
				const reopened = readXlsx(written.value)
				expectOk(reopened)
				expect(reopened.value.workbook.sheets.length).toBe(initial.value.workbook.sheets.length)
			})
		}

		it('keeps fixture set stable', () => {
			const names = fixtureNames().map((name) => basename(name))
			expect(names).toEqual([
				'layout_breaks.xlsx',
				'multisheet_names.xlsx',
				'strings_links.xlsx',
				'styles_formulas.xlsx',
			])
		})
	})
} else {
	describe('XlsxWriter fixture corpus', () => {
		it('is empty when fixtures are unavailable', () => {
			expect(fixtureNames()).toHaveLength(0)
		})
	})
}
