import { describe, expect, test } from 'bun:test'
import {
	packageActionFixtureScanMarkdown,
	runPackageActionFixtureScan,
} from './package-action-fixture-scan.ts'

describe('package action fixture replacement scan', () => {
	test('shows which generated package-action edge cases still lack public candidates', () => {
		const result = runPackageActionFixtureScan()

		expect(result.corpus).toBe('tracked-git-fixtures')
		expect(result.scanned).toBeGreaterThan(200)
		expect(result.skippedDirectories).toEqual(['external', 'stress'])
		expect(result.replacementStatus).toBe('remaining-generated-edge-cases')
		expect(result.featureCounts.docPropsCore).toBeGreaterThan(0)
		expect(result.featureCounts.docPropsCustom).toBeGreaterThan(0)
		expect(result.featureCounts.calcChain).toBeGreaterThan(0)
		expect(result.featureCounts.customXml).toBeGreaterThan(0)
		expect(result.featureCounts.macro).toBeGreaterThan(0)
		expect(result.featureCounts.chartOrDrawing).toBeGreaterThan(0)
		expect(result.featureCounts.signaturePackage).toBe(0)
		expect(result.featureCounts.unknownPathFamily).toBeGreaterThan(0)
		expect(result.rejected).toBe(result.rejectedFixtures.length)
		expect(result.boundary).toContain('tracked public XLSX/XLSM fixture corpus')
	})

	test('renders claim-safe markdown with replacement status and feature counts', () => {
		const markdown = packageActionFixtureScanMarkdown(runPackageActionFixtureScan())

		expect(markdown).toContain('Package Action Fixture Replacement Scan')
		expect(markdown).toContain('Corpus: tracked-git-fixtures')
		expect(markdown).toContain('Skipped directories: external, stress')
		expect(markdown).toContain('Replacement status: remaining-generated-edge-cases')
		expect(markdown).toContain('| signaturePackage | 0 |')
		expect(markdown).toContain('| unknownPathFamily |')
		expect(markdown).toContain('does not authorize hiding generated fixture provenance')
		expect(markdown).toContain('Rejected fixtures:')
	})
})
