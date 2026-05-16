import { describe, expect, test } from 'bun:test'
import { runSafeOpenFixtureScan, safeOpenFixtureScanMarkdown } from './safe-open-fixture-scan.ts'

describe('safe open fixture replacement scan', () => {
	test('shows whether tracked public fixtures can replace synthetic signed and unknown cases', () => {
		const result = runSafeOpenFixtureScan()

		expect(result.corpus).toBe('tracked-git-fixtures')
		expect(result.scanned).toBeGreaterThan(200)
		expect(result.skippedDirectories).toEqual(['external', 'stress'])
		expect(result.replacementStatus).toBe('candidate-found')
		expect(result.signatureOrUnknownMatches.map((entry) => entry.fixture)).toContain(
			'fixtures/xlsx/excelforge/Book_1_unknown_part.xlsx',
		)
		expect(result.riskFamilyCounts.preservedMacro).toBeGreaterThan(0)
		expect(result.riskFamilyCounts.preservedActiveX).toBeGreaterThan(0)
		expect(result.riskFamilyCounts.preservedSignature ?? 0).toBe(0)
		expect(result.riskFamilyCounts.preservedOther).toBeGreaterThan(0)
		expect(result.rejected).toBe(result.rejectedFixtures.length)
		expect(result.boundary).toContain('tracked public XLSX/XLSM fixture corpus')
	})

	test('renders claim-safe markdown with replacement status', () => {
		const markdown = safeOpenFixtureScanMarkdown(runSafeOpenFixtureScan())

		expect(markdown).toContain('Safe Open Fixture Replacement Scan')
		expect(markdown).toContain('Corpus: tracked-git-fixtures')
		expect(markdown).toContain('Skipped directories: external, stress')
		expect(markdown).toContain('Replacement status:')
		expect(markdown).toContain('candidate-found')
		expect(markdown).toContain('fixtures/xlsx/excelforge/Book_1_unknown_part.xlsx')
		expect(markdown).toContain('Risk family counts:')
		expect(markdown).toContain('preservedMacro=')
		expect(markdown).toContain('Rejected fixtures:')
		expect(markdown).toContain('does not prove')
	})
})
