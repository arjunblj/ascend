import { describe, expect, test } from 'bun:test'
import {
	defaultPackageActionProofCases,
	packageActionCompactReleaseReport,
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
		expect(cases.filter((entry) => entry.sourceKind === 'public-fixture')).toHaveLength(4)
		expect(cases.filter((entry) => entry.sourceKind === 'generated-workbook')).toHaveLength(2)
		expect(cases.filter((entry) => entry.sourceKind === 'generated-edge-package')).toHaveLength(2)
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
		expect(proof.cases.every((entry) => /^[a-f0-9]{64}$/.test(entry.inputSha256))).toBe(true)
		expect(proof.cases.every((entry) => entry.commitJournalExact === false)).toBe(true)
		expect(proof.cases.every((entry) => entry.commitJournalPackageIssueCount > 0)).toBe(true)
		expect(proof.cases.every((entry) => entry.commitJournalPackageIssueRefs.length > 0)).toBe(true)
		expect(proof.cases.filter((entry) => entry.sourceKind === 'public-fixture')).toHaveLength(4)
		expect(
			proof.cases.filter((entry) => entry.sourceKind === 'generated-edge-package'),
		).toHaveLength(2)
		expect(proof.cases.find((entry) => entry.name === 'docprops-passthrough')).toMatchObject({
			sourceKind: 'public-fixture',
			fixture: 'fixtures/xlsx/calamine/date_1904.xlsx',
		})
		expect(proof.cases.find((entry) => entry.name === 'calc-chain-drop')).toMatchObject({
			sourceKind: 'public-fixture',
			fixture: 'fixtures/xlsx/poi/Booleans.xlsx',
			postWriteAuditsPassed: true,
			issueCount: 0,
		})
		expect(proof.cases.find((entry) => entry.name === 'unknown-part-error')).toMatchObject({
			postWriteAuditsPassed: false,
			issueCount: 1,
		})
		expect(proof.cases.find((entry) => entry.name === 'docprops-passthrough')).toMatchObject({
			streamingProof: {
				expectedActionsPresent: true,
				streamingRegeneratePartPaths: ['xl/worksheets/sheet1.xml'],
				passthroughBytesEqualCount: expect.any(Number),
				issueCount: 0,
			},
		})
		expect(
			proof.cases.find((entry) => entry.name === 'docprops-passthrough')?.streamingProof
				?.actionCounts.passthrough,
		).toBeGreaterThan(0)
		expect(
			proof.cases.find((entry) => entry.name === 'docprops-passthrough')?.streamingProof
				?.outputByteDigestCount,
		).toBeGreaterThan(0)
	})

	test('renders claim-safe markdown report boundaries', async () => {
		const proof = await runPackageActionProof({ includeTimings: false })
		const markdown = packageActionProofMarkdown(proof)

		expect(markdown).toContain('Package Action Proof Report')
		expect(markdown).toContain('not signed provenance')
		expect(markdown).toContain('Excel recalculation equivalence')
		expect(markdown).toContain('Journal package issues')
		expect(markdown).toContain('Combined commit actions:')
		expect(markdown).toContain('Streaming proof')
		expect(markdown).toContain('streaming regenerate=xl/worksheets/sheet1.xml')
		expect(markdown).toContain('unknown-part-error')
	})

	test('renders compact release report without embedding full proof artifacts', async () => {
		const proof = await runPackageActionProof({ includeTimings: false })
		const compact = packageActionCompactReleaseReport(proof)
		const compactJson = JSON.stringify(compact)
		const fullJson = JSON.stringify(proof)

		expect(compact.claim).toBe('auditable package-part mutation')
		expect(compact.headlineClaimAllowed).toBe(false)
		expect(compact.releaseGate).toBe('blocked-by-publication-policy')
		expect(compact.readyWhen.map((entry) => entry.id)).toContain('streaming-matrix-boundary')
		expect(compact.coverage).toMatchObject({
			cases: 8,
			expectedActionsEverywhere: true,
			sourceGraphEverywhere: true,
			packageJournalIssuesEverywhere: true,
			postWriteAuditFailures: ['unknown-part-error'],
			proofIssueCases: ['unknown-part-error'],
			streamingProofCases: 1,
			streamingRegenerateParts: 1,
		})
		expect(compact.sourceCaseCounts).toEqual({
			'public-fixture': 4,
			'generated-workbook': 2,
			'generated-edge-package': 2,
		})
		expect(compactJson.length).toBeLessThan(fullJson.length)
		expect(compactJson).not.toContain('inputSha256')
		expect(compactJson).not.toContain('outputBytes')
		expect(compactJson).not.toContain('proofJsonBytes')
		expect(compactJson).not.toContain('streamingRegeneratePartPaths')
		expect(compact.boundary).toContain('not signed provenance')
		expect(compact.boundary).toContain('full streaming parity')
	})
})
