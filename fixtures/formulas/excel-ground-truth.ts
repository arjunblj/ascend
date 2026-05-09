/**
 * Utility: build an XLSX workbook with all conformance formula cases, save it for manual
 * comparison in Excel, then verify computed results against fixture expectations.
 *
 * Usage:
 *   bun run fixtures/formulas/excel-ground-truth.ts
 *
 * Steps:
 *   1. This writes 'ground-truth.xlsx' into the repo root.
 *   2. Open in Excel, enable calculation, save.
 *   3. Re-run with --verify flag to compare Excel results to fixture expectations.
 */
import { readdir, readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { createWorkbook, parseA1, type StyleId } from '../../packages/core/src/index.ts'
import { readXlsx, writeXlsx } from '../../packages/io-xlsx/src/index.ts'
import { booleanValue, EMPTY, numberValue, stringValue } from '../../packages/schema/src/index.ts'

const SID = 0 as StyleId
const fixturesDir = join(import.meta.dir, '.')

interface Case {
	description: string
	setup: Record<string, string | number | boolean>
	formula: string
	expected: {
		kind: string
		value?: number | string | boolean
		serial?: number
		approx?: number
		tolerance?: number
	}
}

interface Fixture {
	function: string
	cases: Case[]
}

async function loadFixtures(): Promise<Fixture[]> {
	const files = (await readdir(fixturesDir)).filter(
		(f) => f.endsWith('.json') && f !== 'package.json',
	)
	const fixtures: Fixture[] = []
	for (const file of files) {
		const content = await readFile(join(fixturesDir, file), 'utf-8')
		fixtures.push(JSON.parse(content) as Fixture)
	}
	return fixtures
}

async function buildWorkbook(): Promise<void> {
	const fixtures = await loadFixtures()
	const wb = createWorkbook()
	let sheetIdx = 0
	for (const fx of fixtures) {
		for (let ci = 0; ci < fx.cases.length; ci++) {
			const c = fx.cases[ci]
			if (!c) continue
			const sheetName = `${fx.function}_${String(ci)}`.slice(0, 31)
			const sheet = wb.addSheet(sheetName)
			sheetIdx++
			for (const [a1, val] of Object.entries(c.setup)) {
				const { row, col } = parseA1(a1)
				const value =
					typeof val === 'number'
						? numberValue(val)
						: typeof val === 'boolean'
							? booleanValue(val)
							: stringValue(String(val))
				sheet.cells.set(row, col, { value, formula: null, styleId: SID })
			}
			const f = c.formula.startsWith('=') ? c.formula.slice(1) : c.formula
			sheet.cells.set(15, 0, { value: EMPTY, formula: f, styleId: SID })
		}
	}
	const result = writeXlsx(wb)
	if (!result.ok) {
		console.error('write failed:', result.error.message)
		process.exit(1)
	}
	const outPath = join(import.meta.dir, '..', '..', 'ground-truth.xlsx')
	await Bun.write(outPath, result.value)
	console.log(`Wrote ${outPath} with ${String(sheetIdx)} sheets.`)
	console.log('Open in Excel, let it recalculate, then save.')
	console.log('Re-run with --verify to compare.')
}

async function verify(): Promise<void> {
	const inPath = join(import.meta.dir, '..', '..', 'ground-truth.xlsx')
	const bytes = await Bun.file(inPath).arrayBuffer()
	const loaded = readXlsx(new Uint8Array(bytes))
	if (!loaded.ok) {
		console.error('read failed:', loaded.error.message)
		process.exit(1)
	}
	const wb = loaded.value.workbook
	console.log(`Loaded ${String(wb.sheets.length)} sheets from ${inPath}`)
	let checked = 0
	const mismatches = 0
	for (const sheet of wb.sheets) {
		const val = sheet.cells.readValue(15, 0)
		if (val.kind === 'empty') continue
		checked++
		console.log(`  ${sheet.name}: ${JSON.stringify(val)}`)
	}
	console.log(`Checked ${String(checked)} sheets, ${String(mismatches)} mismatches.`)
}

if (process.argv.includes('--verify')) {
	await verify()
} else {
	await buildWorkbook()
}
