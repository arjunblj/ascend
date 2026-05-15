import { describe, expect, test } from 'bun:test'
import {
	defaultPackageActionProofCases,
	packageActionProofMarkdown,
	runPackageActionProof,
} from './package-action-proof.ts'

describe('package action proof harness', () => {
	test('covers release-claim package action scenarios', () => {
		const cases = defaultPackageActionProofCases()

		expect(cases.map((entry) => entry.name)).toEqual([
			'docprops-passthrough',
			'regenerate-existing-sheet',
			'add-sheet-part',
			'calc-chain-drop',
			'signature-invalidation-drop',
			'macro-passthrough',
			'chart-sidecar-accounting',
			'unknown-part-error',
		])
		expect(cases.filter((entry) => entry.sourceKind === 'public-fixture')).toHaveLength(2)
		expect(cases.filter((entry) => entry.sourceKind === 'generated-workbook')).toHaveLength(2)
		expect(cases.filter((entry) => entry.sourceKind === 'generated-edge-package')).toHaveLength(4)
	})

	test('proves all package action kinds without relying on timing thresholds', async () => {
		const proof = await runPackageActionProof({ includeTimings: false })

		expect(proof.combinedCommitActionCounts.passthrough).toBeGreaterThan(0)
		expect(proof.combinedCommitActionCounts.regenerate).toBeGreaterThan(0)
		expect(proof.combinedCommitActionCounts.add).toBeGreaterThan(0)
		expect(proof.combinedCommitActionCounts.drop).toBeGreaterThan(0)
		expect(proof.combinedCommitActionCounts.error).toBeGreaterThan(0)
		expect(proof.cases.every((entry) => entry.expectedActionsPresent)).toBe(true)
		expect(proof.cases.every((entry) => entry.commitCoverage.sourceGraphIncluded)).toBe(true)
		expect(proof.cases.every((entry) => entry.commitCoverage.outputByteDigestCount > 0)).toBe(true)
		expect(proof.cases.every((entry) => entry.commitJournalExact === false)).toBe(true)
		expect(proof.cases.every((entry) => entry.commitJournalPackageIssueCount > 0)).toBe(true)
		expect(proof.cases.every((entry) => entry.commitJournalPackageIssueRefs.length > 0)).toBe(true)
		expect(proof.cases.filter((entry) => entry.sourceKind === 'public-fixture')).toHaveLength(2)
		expect(
			proof.cases.filter((entry) => entry.sourceKind === 'generated-edge-package'),
		).toHaveLength(4)
		expect(proof.cases.find((entry) => entry.name === 'unknown-part-error')).toMatchObject({
			postWriteAuditsPassed: false,
			issueCount: 1,
		})
	})

	test('renders claim-safe markdown report boundaries', async () => {
		const proof = await runPackageActionProof({ includeTimings: false })
		const markdown = packageActionProofMarkdown(proof)

		expect(markdown).toContain('Package Action Proof Report')
		expect(markdown).toContain('not signed provenance')
		expect(markdown).toContain('Excel recalculation equivalence')
		expect(markdown).toContain('Journal package issues')
		expect(markdown).toContain('Combined commit actions:')
		expect(markdown).toContain('unknown-part-error')
	})
})
