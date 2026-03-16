import { describe, expect, it } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { readXlsx } from '../../../packages/io-xlsx/src/index.ts'

const STRESS_DIR = join(dirname(fileURLToPath(import.meta.url)), '.')

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
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

function countRows(workbook: {
	sheets: readonly {
		name?: string
		cells: { iterate: () => Iterable<[number, number, unknown]> }
	}[]
}): number {
	let totalRows = 0
	for (const sheet of workbook.sheets) {
		let maxRow = -1
		for (const [row] of sheet.cells.iterate()) {
			if (row > maxRow) maxRow = row
		}
		totalRows += maxRow + 1
	}
	return totalRows
}

const FIXTURES = [
	{ name: 'dense-100k.xlsx', minRows: 100_000, minCells: 500_000, minSheets: 1 },
	{ name: 'many-styles.xlsx', minRows: 1000, minCells: 2000, minSheets: 1 },
	{ name: 'many-strings.xlsx', minRows: 10_000, minCells: 20_000, minSheets: 1 },
	{ name: 'formula-dense.xlsx', minRows: 5_000, minCells: 25_000, minSheets: 1 },
	{ name: 'merged-complex.xlsx', minRows: 40, minCells: 200, minSheets: 1 },
	{ name: 'multi-sheet-10.xlsx', minRows: 10_000, minCells: 10_000, minSheets: 10 },
] as const

function ensureFixtures(): void {
	if (FIXTURES.every((f) => existsSync(join(STRESS_DIR, f.name)))) return
	const result = Bun.spawnSync({
		cmd: ['bun', 'run', join(dirname(STRESS_DIR), 'generate-stress.ts')],
		cwd: dirname(STRESS_DIR),
		stdout: 'pipe',
		stderr: 'pipe',
	})
	if (result.exitCode !== 0) {
		throw new Error(
			`Failed to generate stress fixtures: ${new TextDecoder().decode(result.stderr || new Uint8Array())}`,
		)
	}
}

describe('Stress XLSX fixtures', () => {
	ensureFixtures()
	for (const fixture of FIXTURES) {
		const path = join(STRESS_DIR, fixture.name)
		it(`reads ${fixture.name} without crashing`, () => {
			const bytes = readFileSync(path)
			const result = readXlsx(new Uint8Array(bytes))
			expectOk(result)
			expect(result.value.workbook.sheets.length).toBeGreaterThan(0)
		})

		it(`verifies basic counts for ${fixture.name}`, () => {
			const bytes = readFileSync(path)
			const result = readXlsx(new Uint8Array(bytes))
			expectOk(result)
			const wb = result.value.workbook
			expect(wb.sheets.length).toBeGreaterThanOrEqual(fixture.minSheets)
			const cellCount = countCells(wb)
			expect(cellCount).toBeGreaterThanOrEqual(fixture.minCells)
			const totalRows = countRows(wb)
			expect(totalRows).toBeGreaterThanOrEqual(fixture.minRows)
		})
	}
})
