import { describe, expect, test } from 'bun:test'
import {
	agentViewRecoveryProofMarkdown,
	runAgentViewRecoveryProof,
} from './agent-view-recovery-proof.ts'

describe('agent view omitted evidence recovery proof', () => {
	test('proves same-range unbudgeted recovery and count-only budget metadata', async () => {
		const proof = await runAgentViewRecoveryProof()

		expect(proof.cases).toHaveLength(5)
		expect(proof.allUnbudgetedRecoveriesExact).toBe(true)
		expect(proof.allBudgetMetadataCountOnly).toBe(true)
		expect(proof.narrowerRecoveryRequiresLocatorMetadata).toBe(true)
		expect(proof.cases.some((entry) => !entry.withinBudget)).toBe(true)
		expect(
			proof.cases.every(
				(entry) =>
					entry.recoveredSampleRows === entry.omittedSampleRows &&
					entry.recoveredColumnSampleValues === entry.omittedColumnSampleValues &&
					entry.recoveredFormulaPatterns === entry.omittedFormulaPatterns,
			),
		).toBe(true)
	})

	test('renders recovery boundaries without promoting automatic narrow reads', async () => {
		const markdown = agentViewRecoveryProofMarkdown(await runAgentViewRecoveryProof())

		expect(markdown).toContain('Agent View Omitted Evidence Recovery Proof')
		expect(markdown).toContain('omission counts, not omitted row/column/formula locations')
		expect(markdown).toContain('Same-range unbudgeted recovery is exact')
		expect(markdown).toContain('Narrower recovery requires locator metadata: true')
	})
})
