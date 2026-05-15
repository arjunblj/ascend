import { describe, expect, test } from 'bun:test'
import {
	agentViewRecoveryProofMarkdown,
	runAgentViewRecoveryProof,
} from './agent-view-recovery-proof.ts'

describe('agent view omitted evidence recovery proof', () => {
	test('proves same-range unbudgeted recovery and compact locator metadata', async () => {
		const proof = await runAgentViewRecoveryProof()

		expect(proof.cases).toHaveLength(5)
		expect(proof.allUnbudgetedRecoveriesExact).toBe(true)
		expect(proof.allBudgetMetadataHasLocators).toBe(true)
		expect(proof.allSampleRowLocatorsExact).toBe(true)
		expect(proof.allColumnSampleLocatorsExact).toBe(true)
		expect(proof.allNarrowSampleRowRecoveriesExact).toBe(true)
		expect(proof.allFormulaPatternExampleRecoveriesExact).toBe(true)
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
		expect(markdown).toContain('compact omitted sample-row and column-sample locators')
		expect(markdown).toContain('Same-range unbudgeted recovery is exact')
		expect(markdown).toContain('All budget metadata has locators: true')
		expect(markdown).toContain('All narrow sample-row recoveries exact: true')
		expect(markdown).toContain('All formula-pattern example recoveries exact: true')
	})
})
