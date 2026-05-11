import { describe, expect, test } from 'bun:test'
import { runHyperFormulaComparison } from './formula-hyperformula-compare.ts'

describe('formula correctness vs HyperFormula', () => {
	test(
		'matches the shared public OSS formula scenario set',
		() => {
			const result = runHyperFormulaComparison()

			expect(result.scenarios).toBeGreaterThanOrEqual(700)
			expect(result.rows.filter((row) => row.status === 'match').length).toBeGreaterThanOrEqual(450)
			expect(result.knownDivergences).toBeGreaterThanOrEqual(30)
			expect(result.mismatches).toBe(0)
			expect(result.rows.filter((row) => row.status === 'diff')).toEqual([])
		},
		{ timeout: 30_000 },
	)
})
