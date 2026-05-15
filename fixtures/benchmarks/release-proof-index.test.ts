import { describe, expect, test } from 'bun:test'
import { releaseProofIndexMarkdown, runReleaseProofIndex } from './release-proof-index.ts'

describe('release proof evidence index', () => {
	test('references top claim proof artifacts by digest without embedding artifacts', async () => {
		const index = await runReleaseProofIndex({ includeTimings: false })

		expect(index.signed).toBe(false)
		expect(index.attestation).toBe(false)
		expect(index.excludedEvidenceCount).toBe(1)
		expect(index.artifacts.map((artifact) => artifact.name)).toEqual([
			'safe-open-proof',
			'package-action-proof',
		])
		expect(index.excludedEvidence).toEqual([
			expect.objectContaining({
				name: 'practical-latency-contracts',
				ownerLoop: 'performance',
				eligibilityRule: expect.stringContaining('tracked-clean'),
			}),
		])
		for (const artifact of index.artifacts) {
			expect(artifact.sha256).toMatch(/^[a-f0-9]{64}$/)
			expect(artifact.stableShapeSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(artifact.jsonBytes).toBeGreaterThan(100)
			expect(artifact.markdownBytes).toBeGreaterThan(100)
		}
		expect(index.artifacts.find((artifact) => artifact.name === 'safe-open-proof')).toMatchObject({
			command: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json',
			publicationStatus: 'needs-release-packaging',
			headlineClaimAllowed: false,
			releaseGate: 'blocked-by-publication-policy',
			fixtureProvenance: {
				publicFixtureCases: 6,
				generatedWorkbookCases: 0,
				generatedEdgePackageCases: 2,
				malformedCases: 1,
				generatedCaseNames: ['signed', 'unknown-part', 'malformed'],
			},
			summary: { cases: 9, rejected: 1, malformedRejected: true },
		})
		expect(
			index.artifacts
				.find((artifact) => artifact.name === 'safe-open-proof')
				?.publicationBlockers.join(' '),
		).toContain('public binary fixtures')
		expect(
			index.artifacts.find((artifact) => artifact.name === 'package-action-proof'),
		).toMatchObject({
			command: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
			publicationStatus: 'needs-release-packaging',
			headlineClaimAllowed: false,
			releaseGate: 'blocked-by-publication-policy',
			fixtureProvenance: {
				publicFixtureCases: 2,
				generatedWorkbookCases: 2,
				generatedEdgePackageCases: 4,
				malformedCases: 0,
			},
			summary: { cases: 8, allActionsCovered: true, sourceGraphEverywhere: true },
		})
		expect(index.artifacts.every((artifact) => artifact.headlineClaimAllowed === false)).toBe(true)
	})

	test('keeps stable shape digests deterministic in no-timings mode', async () => {
		const first = await runReleaseProofIndex({ includeTimings: false })
		const second = await runReleaseProofIndex({ includeTimings: false })

		expect(first.artifacts.map((artifact) => artifact.stableShapeSha256)).toEqual(
			second.artifacts.map((artifact) => artifact.stableShapeSha256),
		)
	})

	test('renders honest non-attestation boundaries', async () => {
		const markdown = releaseProofIndexMarkdown(
			await runReleaseProofIndex({ includeTimings: false }),
		)

		expect(markdown).toContain('Release Proof Evidence Index')
		expect(markdown).toContain('not signed provenance')
		expect(markdown).toContain('Publication blockers')
		expect(markdown).toContain('Headline claim allowed')
		expect(markdown).toContain('blocked-by-publication-policy')
		expect(markdown).toContain('Fixture provenance')
		expect(markdown).toContain('generatedCases=signed,unknown-part,malformed')
		expect(markdown).toContain('safe-open-proof.ts --no-timings --json')
		expect(markdown).toContain('SLSA')
		expect(markdown).toContain('Attestation: false')
		expect(markdown).toContain('safe unknown workbook opening')
		expect(markdown).toContain('auditable package-part mutation')
		expect(markdown).toContain('Excluded Evidence')
		expect(markdown).toContain('practical-latency-contracts')
		expect(markdown).toContain('tracked-clean run')
	})
})
