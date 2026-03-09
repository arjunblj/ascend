import { describe, expect, test } from 'bun:test'
import { readdir, readFile } from 'node:fs/promises'
import { join } from 'node:path'
import type { StyleId } from '@ascend/core'
import { createWorkbook, parseA1 } from '@ascend/core'
import type { CalcContext } from '@ascend/engine'
import { defaultCalcContext, recalculate } from '@ascend/engine'
import type { CellValue } from '@ascend/schema'
import { booleanValue, EMPTY, errorValue, numberValue, stringValue } from '@ascend/schema'

const sid = 0 as StyleId

interface ConformanceCase {
	description: string
	setup: Record<string, string | number | boolean>
	formula: string
	context?: {
		dateSystem?: '1900' | '1904'
		now?: string
		today?: string
		randomSeed?: number
		locale?: string
	}
	expected: {
		kind: string
		value?: number | string | boolean
		serial?: number
		approx?: number
		tolerance?: number
	}
}

interface ConformanceFixture {
	function: string
	cases: ConformanceCase[]
}

function inputToCellValue(v: string | number | boolean): CellValue {
	if (typeof v === 'number') return numberValue(v)
	if (typeof v === 'boolean') return booleanValue(v)
	return stringValue(String(v))
}

function expectedToCellValue(e: ConformanceCase['expected']): CellValue {
	switch (e.kind) {
		case 'number':
			return numberValue(e.value as number)
		case 'string':
			return stringValue(e.value as string)
		case 'boolean':
			return booleanValue(e.value as boolean)
		case 'error':
			return errorValue(e.value as string as import('@ascend/schema').ExcelError)
		case 'empty':
			return EMPTY
		case 'date':
			return { kind: 'date', serial: e.serial ?? (e.value as number) }
		default:
			throw new Error(`Unknown expected kind: ${e.kind}`)
	}
}

function cellValuesEqual(actual: CellValue, expected: CellValue): boolean {
	if (actual.kind !== expected.kind) return false
	switch (actual.kind) {
		case 'empty':
			return true
		case 'number':
			return (
				expected.kind === 'number' &&
				(Math.abs(actual.value - expected.value) < 1e-10 || actual.value === expected.value)
			)
		case 'string':
			return expected.kind === 'string' && actual.value === expected.value
		case 'boolean':
			return expected.kind === 'boolean' && actual.value === expected.value
		case 'error':
			return expected.kind === 'error' && actual.value === expected.value
		case 'date':
			return expected.kind === 'date' && actual.serial === expected.serial
		default:
			return false
	}
}

function runCase(
	_fnName: string,
	c: ConformanceCase,
	baseCtx: CalcContext,
): { pass: boolean; actual?: CellValue; error?: string } {
	const wb = createWorkbook()
	const sheet = wb.addSheet('Sheet1')

	for (const [a1, val] of Object.entries(c.setup)) {
		const { row, col } = parseA1(a1)
		sheet.cells.set(row, col, {
			value: inputToCellValue(val),
			formula: null,
			styleId: sid,
		})
	}

	const formula = c.formula.startsWith('=') ? c.formula.slice(1) : c.formula
	const formulaRow = 10
	const formulaCol = 0
	sheet.cells.set(formulaRow, formulaCol, {
		value: EMPTY,
		formula,
		styleId: sid,
	})

	const ctx: CalcContext = {
		...baseCtx,
		...(c.context?.dateSystem ? { dateSystem: c.context.dateSystem } : {}),
		...(c.context?.randomSeed !== undefined ? { randomSeed: c.context.randomSeed } : {}),
		...(c.context?.locale ? { locale: c.context.locale } : {}),
		...(c.context?.now ? { now: new Date(c.context.now) } : {}),
		...(c.context?.today ? { today: new Date(c.context.today) } : {}),
	}
	wb.calcSettings = {
		...wb.calcSettings,
		dateSystem: ctx.dateSystem,
		iterativeCalc: ctx.iterativeCalc,
	}

	const result = recalculate(wb, ctx)
	if (result.errors.length > 0) {
		const err = result.errors[0]
		return { pass: false, error: err?.error.message }
	}

	const cell = sheet.cells.get(formulaRow, formulaCol)
	const actual = cell?.value ?? EMPTY

	if (c.expected.approx !== undefined && c.expected.tolerance !== undefined) {
		const pass =
			actual.kind === 'number' && Math.abs(actual.value - c.expected.approx) <= c.expected.tolerance
		return { pass, actual }
	}

	const expected = expectedToCellValue(c.expected)
	const pass = cellValuesEqual(actual, expected)

	return { pass, actual }
}

const fixturesDir = join(import.meta.dir, '.')
const jsonFiles = (await readdir(fixturesDir)).filter((f) => f.endsWith('.json'))
const fixtures: Array<{ file: string; fixture: ConformanceFixture }> = []
for (const file of jsonFiles) {
	const content = await readFile(join(fixturesDir, file), 'utf-8')
	fixtures.push({ file, fixture: JSON.parse(content) as ConformanceFixture })
}
const ctx = defaultCalcContext()

describe('formula conformance', () => {
	test('fixtures directory has JSON files', () => {
		expect(jsonFiles.length).toBeGreaterThan(0)
	})

	for (const { file, fixture } of fixtures) {
		const cases = fixture.cases ?? []
		describe(`${file} (${fixture.function})`, () => {
			for (const c of cases) {
				test(c.description, () => {
					const { pass, actual, error } = runCase(fixture.function, c, ctx)
					if (error) {
						expect.fail(error)
					}
					expect(
						pass,
						`Expected ${JSON.stringify(c.expected)}, got ${JSON.stringify(actual)}`,
					).toBe(true)
				})
			}
		})
	}
})
