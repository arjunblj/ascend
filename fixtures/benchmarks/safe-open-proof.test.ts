import { describe, expect, test } from 'bun:test'
import {
	defaultSafeOpenProofCases,
	runSafeOpenProof,
	safeOpenProofMarkdown,
} from './safe-open-proof.ts'

describe('safe open proof harness', () => {
	test('covers public, synthetic, and malformed proof cases', () => {
		const cases = defaultSafeOpenProofCases()

		expect(cases.map((entry) => entry.name)).toEqual([
			'clean',
			'formula-heavy',
			'macro',
			'pivot',
			'activex',
			'chart',
			'signed',
			'unknown-part',
			'malformed',
		])
		expect(cases.filter((entry) => entry.kind === 'file')).toHaveLength(6)
		expect(cases.filter((entry) => entry.kind === 'synthetic')).toHaveLength(2)
		expect(cases.filter((entry) => entry.kind === 'malformed')).toHaveLength(1)
	})

	test('proves routing decisions without relying on timing thresholds', async () => {
		const proof = await runSafeOpenProof({ repeat: 1, warmup: 0, includeTimings: false })
		const byName = new Map(proof.cases.map((entry) => [entry.name, entry]))

		expect(byName.get('clean')).toMatchObject({
			status: 'ok',
			recommendedMode: 'formula',
			reviewBeforeHydration: false,
		})
		expect(byName.get('clean')?.inputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(byName.get('macro')).toMatchObject({
			status: 'ok',
			recommendedMode: 'metadata-only',
			reviewBeforeHydration: true,
			riskFamilies: ['preservedMacro'],
		})
		expect(byName.get('signed')).toMatchObject({
			status: 'ok',
			recommendedMode: 'metadata-only',
			reviewBeforeHydration: true,
			riskFamilies: ['preservedSignature'],
		})
		expect(byName.get('unknown-part')).toMatchObject({
			status: 'ok',
			recommendedMode: 'metadata-only',
			reviewBeforeHydration: true,
			riskFamilies: ['preservedOther'],
		})
		expect(byName.get('malformed')).toMatchObject({
			status: 'rejected',
		})
		expect(byName.get('malformed')?.boundary).toContain('open-plan rejected:')
	})

	test('renders claim-safe markdown report boundaries', async () => {
		const proof = await runSafeOpenProof({ repeat: 1, warmup: 0, includeTimings: false })
		const markdown = safeOpenProofMarkdown(proof)

		expect(markdown).toContain('Safe Unknown Workbook Opening Proof')
		expect(markdown).toContain('pre-hydration package-feature routing')
		expect(markdown).toContain('not malware scanning')
		expect(markdown).toContain('Allowed claim:')
		expect(markdown).toContain('metadata-only')
	})
})
