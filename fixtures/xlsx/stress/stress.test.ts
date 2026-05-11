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

function countSheetStats(workbook: {
	sheets: readonly {
		cells: { iterate: () => Iterable<[number, number, unknown]> }
	}[]
}): { readonly cells: number; readonly rows: number } {
	let cells = 0
	let rows = 0
	for (const sheet of workbook.sheets) {
		let maxRow = -1
		for (const [row] of sheet.cells.iterate()) {
			cells++
			if (row > maxRow) maxRow = row
		}
		rows += maxRow + 1
	}
	return { cells, rows }
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
		it(`reads and verifies basic counts for ${fixture.name}`, () => {
			const bytes = readFileSync(path)
			const result = readXlsx(new Uint8Array(bytes))
			expectOk(result)
			const wb = result.value.workbook
			expect(wb.sheets.length).toBeGreaterThanOrEqual(fixture.minSheets)
			const stats = countSheetStats(wb)
			expect(stats.cells).toBeGreaterThanOrEqual(fixture.minCells)
			expect(stats.rows).toBeGreaterThanOrEqual(fixture.minRows)
		})
	}
})
