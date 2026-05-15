import { describe, expect, test } from 'bun:test'
import { runSafeOpenFixtureScan, safeOpenFixtureScanMarkdown } from './safe-open-fixture-scan.ts'

describe('safe open fixture replacement scan', () => {
	test('shows whether checked-in public fixtures can replace synthetic signed and unknown cases', () => {
		const result = runSafeOpenFixtureScan()

		expect(result.scanned).toBeGreaterThan(300)
		expect(result.replacementStatus).toBe('no-public-binary-replacement-found')
		expect(result.signatureOrUnknownMatches).toHaveLength(0)
		expect(result.rejected).toBe(result.rejectedFixtures.length)
		expect(result.boundary).toContain('checked-in public XLSX/XLSM fixtures only')
	})

	test('renders claim-safe markdown with replacement status', () => {
		const markdown = safeOpenFixtureScanMarkdown(runSafeOpenFixtureScan())

		expect(markdown).toContain('Safe Open Fixture Replacement Scan')
		expect(markdown).toContain('Replacement status:')
		expect(markdown).toContain('no-public-binary-replacement-found')
		expect(markdown).toContain('Rejected fixtures:')
		expect(markdown).toContain('does not prove')
	})
})
