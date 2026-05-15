import { describe, expect, test } from 'bun:test'
import {
	agentViewBudgetProofMarkdown,
	defaultAgentViewBudgetProofCases,
	runAgentViewBudgetProof,
} from './agent-view-budget-proof.ts'

describe('agent view budget proof harness', () => {
	test('covers mixed workbook shapes', () => {
		const cases = defaultAgentViewBudgetProofCases()

		expect(cases.map((entry) => entry.name)).toEqual([
			'dense-table',
			'wide-sparse',
			'formula-heavy',
			'metadata-heavy',
			'public-formula-stress',
		])
	})

	test('proves deterministic budget metadata and shape preservation', async () => {
		const proof = await runAgentViewBudgetProof()

		expect(proof.allDeterministic).toBe(true)
		expect(proof.allShapePreserved).toBe(true)
		expect(proof.allOmissionsCounted).toBe(true)
		expect(
			proof.cases.every((entry) => entry.unbudgetedApproxTokens === entry.fullApproxTokens),
		).toBe(true)
		expect(proof.cases.every((entry) => entry.compressionRatio < 1)).toBe(true)
		expect(proof.cases.some((entry) => !entry.withinBudget)).toBe(true)
		expect(proof.cases.find((entry) => entry.name === 'public-formula-stress')).toMatchObject({
			omittedFormulaPatterns: 6,
			withinBudget: true,
		})
	})

	test('renders claim-safe markdown boundaries', async () => {
		const proof = await runAgentViewBudgetProof()
		const markdown = agentViewBudgetProofMarkdown(proof)

		expect(markdown).toContain('Agent View Budget Proof')
		expect(markdown).toContain('not model-specific tokenization')
		expect(markdown).toContain('Omitted evidence is intentionally absent')
		expect(markdown).toContain('All truncations counted: true')
	})
})
