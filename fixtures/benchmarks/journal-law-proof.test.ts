import { describe, expect, test } from 'bun:test'
import {
	journalLawClaimReport,
	journalLawClaimReportMarkdown,
	journalLawProofMarkdown,
	runJournalLawProof,
} from './journal-law-proof.ts'

describe('journal law proof harness', () => {
	test('covers generated exact inverse laws and lossy metadata boundaries', () => {
		const result = runJournalLawProof({ exactCaseCount: 48, sequenceLength: 5 })

		expect(result.passed).toBe(true)
		expect(result.failureCount).toBe(0)
		expect(result.exactChecked).toBe(56)
		expect(result.lossyChecked).toBe(7)
		expect(Object.keys(result.operationFamilies).length).toBeGreaterThanOrEqual(24)
		expect(result.operationFamilies.setDataValidation ?? 0).toBeGreaterThan(0)
		expect(result.operationFamilies.setConditionalFormat ?? 0).toBeGreaterThan(0)
		expect(result.operationFamilies.setRowHeight ?? 0).toBeGreaterThan(0)
		expect(result.operationFamilies.setColWidth ?? 0).toBeGreaterThan(0)
		expect(result.operationFamilies.setPageSetup ?? 0).toBeGreaterThan(0)
		expect(result.operationFamilies.setDocumentProperties ?? 0).toBeGreaterThan(0)
		expect(result.operationFamilies.setTheme ?? 0).toBeGreaterThan(0)
		expect(result.operationFamilies.setTableStyle ?? 0).toBeGreaterThan(0)
		expect(
			result.cases.find((entry) => entry.name === 'theme-replacement')?.operationNames,
		).toEqual(['setTheme'])
		expect(result.issueReasons['package-parts:package-part-preservation']).toBeGreaterThan(0)
		expect(result.issueReasons['tables:table-metadata']).toBe(1)
		expect(result.issueReasons['data-validations:metadata-order']).toBe(1)
		expect(result.issueReasons['data-validations:metadata-duplicate']).toBe(1)
		expect(result.issueReasons['conditional-formats:metadata-order']).toBe(2)
		expect(result.issueReasons['conditional-formats:metadata-duplicate']).toBe(1)
		expect(result.cases.every((entry) => entry.passed)).toBe(true)
	})

	test('renders claim-safe proof boundaries', () => {
		const markdown = journalLawProofMarkdown(
			runJournalLawProof({ exactCaseCount: 4, sequenceLength: 3 }),
		)

		expect(markdown).toContain('deterministic generation exercises journal inverse laws')
		expect(markdown).toContain('not shrinkable property testing')
		expect(markdown).toContain('data-validations:metadata-order')
		expect(markdown).toContain('conditional-formats:metadata-duplicate')
	})

	test('renders release-facing claim wording without promoting lossy boundaries', () => {
		const report = journalLawClaimReport(
			runJournalLawProof({ exactCaseCount: 4, sequenceLength: 3 }),
		)
		const markdown = journalLawClaimReportMarkdown(report)

		expect(report.proofStatus).toBe('passed')
		expect(report.exactOperationFamilies).toContain('setTheme')
		expect(report.lossyIssueReasons['package-parts:package-part-preservation']).toBeGreaterThan(0)
		expect(markdown).toContain('Claim wording allowed today')
		expect(markdown).toContain('not shrinkable property-based testing')
		expect(markdown).toContain('Style and table-style exactness')
	})
})
