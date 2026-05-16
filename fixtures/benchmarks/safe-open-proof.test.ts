import { describe, expect, test } from 'bun:test'
import {
	defaultSafeOpenProofCases,
	runSafeOpenProof,
	safeOpenCompactReleaseReport,
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
			'encrypted-password',
			'encrypted-missing-password',
			'encrypted-wrong-password',
			'malformed',
		])
		expect(cases.filter((entry) => entry.kind === 'file')).toHaveLength(10)
		expect(cases.filter((entry) => entry.kind === 'synthetic')).toHaveLength(1)
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
			kind: 'file',
			fixture: 'fixtures/xlsx/excelforge/Book_1_unknown_part.xlsx',
			status: 'ok',
			recommendedMode: 'metadata-only',
			reviewBeforeHydration: true,
			riskFamilies: ['preservedOther'],
		})
		expect(byName.get('malformed')).toMatchObject({
			status: 'rejected',
		})
		expect(byName.get('encrypted-password')).toMatchObject({
			kind: 'file',
			fixture: 'fixtures/xlsx/calamine/pass_protected.xlsx',
			status: 'ok',
			recommendedMode: 'full',
			reviewBeforeHydration: false,
		})
		expect(byName.get('encrypted-missing-password')).toMatchObject({
			kind: 'file',
			fixture: 'fixtures/xlsx/calamine/pass_protected.xlsx',
			status: 'rejected',
		})
		expect(byName.get('encrypted-wrong-password')).toMatchObject({
			kind: 'file',
			fixture: 'fixtures/xlsx/calamine/pass_protected.xlsx',
			status: 'rejected',
		})
		expect(byName.get('encrypted-missing-password')?.boundary).toContain('requires a password')
		expect(byName.get('encrypted-wrong-password')?.boundary).toContain('Invalid XLSX password')
		expect(byName.get('malformed')?.boundary).toContain('open-plan rejected:')
	})

	test('renders claim-safe markdown report boundaries', async () => {
		const proof = await runSafeOpenProof({ repeat: 1, warmup: 0, includeTimings: false })
		const markdown = safeOpenProofMarkdown(proof)

		expect(proof.timingEnvironment).toBeUndefined()
		expect(markdown).toContain('Safe Unknown Workbook Opening Proof')
		expect(markdown).toContain('pre-hydration package-feature routing')
		expect(markdown).toContain('not malware scanning')
		expect(markdown).toContain('Timing environment: not captured')
		expect(markdown).toContain('Allowed claim:')
		expect(markdown).toContain('metadata-only')
	})

	test('reports latency distribution fields without turning them into thresholds', async () => {
		const proof = await runSafeOpenProof({ repeat: 2, warmup: 0, includeTimings: true })
		const clean = proof.cases.find((entry) => entry.name === 'clean')
		const markdown = safeOpenProofMarkdown(proof)

		expect(clean).toMatchObject({
			openPlanSampleCount: 2,
			fullOpenSampleCount: 2,
		})
		expect(proof.timingEnvironment).toMatchObject({
			runtime: 'bun',
			bunVersion: expect.any(String),
			nodeVersion: expect.stringMatching(/^v/),
			platform: expect.any(String),
			arch: expect.any(String),
			cpuModel: expect.any(String),
			cpuCount: expect.any(Number),
			totalMemoryBytes: expect.any(Number),
			boundary: expect.stringContaining('owner review only'),
		})
		expect(clean?.openPlanMedianMs).toBeGreaterThan(0)
		expect(clean?.openPlanP95Ms).toBeGreaterThan(0)
		expect(clean?.openPlanCv).toBeGreaterThanOrEqual(0)
		expect(clean?.fullOpenMedianMs).toBeGreaterThan(0)
		expect(clean?.fullOpenP95Ms).toBeGreaterThan(0)
		expect(clean?.fullOpenCv).toBeGreaterThanOrEqual(0)
		expect(markdown).toContain('P95 open-plan ms')
		expect(markdown).toContain('Open-plan CV')
		expect(markdown).toContain('Timing environment: runtime=bun')
		expect(markdown).toContain('not malware scanning')
	})

	test('renders compact release report without weakening publication blockers', async () => {
		const proof = await runSafeOpenProof({ repeat: 1, warmup: 0, includeTimings: false })
		const compact = safeOpenCompactReleaseReport(proof)
		const compactJson = JSON.stringify(compact)

		expect(compact.claim).toBe('safe unknown workbook opening')
		expect(compact.headlineClaimAllowed).toBe(false)
		expect(compact.releaseGate).toBe('blocked-by-publication-policy')
		expect(compact.readyWhen.map((entry) => entry.id)).toEqual([
			'public-edge-fixtures',
			'release-latency-run',
			'publication-boundary',
		])
		expect(compact.coverage).toMatchObject({
			cases: 12,
			ok: 9,
			rejected: 3,
			reviewBeforeHydration: 4,
			malformedRejected: true,
			recommendedModes: {
				formula: 4,
				'metadata-only': 4,
				full: 1,
			},
			riskFamilies: ['preservedActiveX', 'preservedMacro', 'preservedOther', 'preservedSignature'],
		})
		expect(compact.caseKindCounts).toEqual({
			file: 10,
			synthetic: 1,
			malformed: 1,
		})
		expect(compactJson).not.toContain('inputSha256')
		expect(compactJson).not.toContain('"bytes":')
		expect(compactJson).not.toContain('"password"')
		expect(compactJson).not.toContain('123')
		expect(compactJson).not.toContain('openPlanMedianMs')
		expect(compact.boundary).toContain('not malware scanning')
		expect(compact.boundary).toContain('release performance threshold')
	})
})
