import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
	emitFormulaLogicTest,
	type FormulaLogicConformanceFixture,
	parseFormulaLogicTest,
	runFormulaLogicTest,
} from './formula-logictest.ts'

const smokeFixture = JSON.parse(
	readFileSync(join(import.meta.dir, 'conformance-smoke.json'), 'utf-8'),
) as FormulaLogicConformanceFixture

describe('formula-logictest conformance harness', () => {
	test('emits and validates completed formula logic tests from JSON fixtures', () => {
		const fixture = {
			function: 'smoke-subset',
			cases: smokeFixture.cases.slice(0, 3),
		}
		const script = emitFormulaLogicTest(fixture, { source: 'conformance-smoke.json' })
		expect(script).toContain('query value label=smoke-subset-areas-returns-1-for-single-range')
		expect(script).toContain('----')
		expect(script).toContain('number approx 0.07177 tolerance 0.001')

		const records = parseFormulaLogicTest(script)
		expect(records).toHaveLength(3)
		expect(records[0]).toMatchObject({
			label: 'smoke-subset-areas-returns-1-for-single-range',
			formula: 'AREAS(A1:B2)',
			expected: { kind: 'number', value: 1 },
		})

		const results = runFormulaLogicTest(script)
		expect(results).toHaveLength(3)
		expect(results.every((result) => result.pass)).toBe(true)
	})
})
