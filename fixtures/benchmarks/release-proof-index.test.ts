import { describe, expect, test } from 'bun:test'
import {
	releaseProofIndexMarkdown,
	releaseProofOwnerHandoffIndex,
	runReleaseProofIndex,
} from './release-proof-index.ts'

describe('release proof evidence index', () => {
	test('references top claim proof artifacts by digest without embedding artifacts', async () => {
		const index = await runReleaseProofIndex({ includeTimings: false })

		expect(index.signed).toBe(false)
		expect(index.attestation).toBe(false)
		expect(index.excludedEvidenceCount).toBe(1)
		expect(index.deferredClaimCount).toBe(6)
		expect(index.fixturePolicy).toMatchObject({
			currentDecision: 'owner-approval-required',
			trackedFixtureScanCommands: {
				'safe-open-proof': 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
				'package-action-proof': 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
			},
			currentGeneratedStructuralCases: {
				'safe-open-proof': ['signed', 'unknown-part', 'malformed'],
				'package-action-proof': ['signature-invalidation-drop', 'unknown-part-error'],
			},
			boundary: expect.stringContaining('not approval to publish generated fixtures'),
		})
		expect(index.fixturePolicy.generatedStructuralFixturesAllowedWhen).toEqual([
			expect.stringContaining('package-topology-only'),
			expect.stringContaining('tracked harness code'),
			expect.stringContaining('tracked public fixture scan'),
			expect.stringContaining('owner gate missing'),
			expect.stringContaining('signed provenance'),
		])
		expect(index.fixturePolicy.publicBinaryFixturesRequiredWhen).toEqual([
			expect.stringContaining('real-world authoring behavior'),
			expect.stringContaining('active-content safety'),
			expect.stringContaining('licensing, privacy, provenance'),
		])
		expect(index.fixturePolicy.sourceReferences.map((entry) => entry.label)).toEqual([
			'GitHub repository limits',
			'GitHub large files',
			'OpenSSF Scorecard binary artifacts',
			'SLSA provenance',
			'GitHub artifact attestations',
		])
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
		expect(index.deferredClaims.map((claim) => claim.name)).toEqual([
			'formula-language-service-primitives',
			'token-bounded-agent-view',
			'retained-viewport-patch-history',
			'columnar-scan-sidecars',
			'formula-oracle-routing',
			'agent-workflow-observability',
		])
		expect(
			index.deferredClaims.find((claim) => claim.name === 'formula-language-service-primitives'),
		).toMatchObject({
			status: 'proof-backed-hold',
			ownerLoops: ['product', 'correctness'],
			reason: expect.stringContaining('edit-producing rename is frozen'),
			killCriterion: expect.stringContaining('Do not promote rename'),
		})
		expect(
			index.deferredClaims.find((claim) => claim.name === 'columnar-scan-sidecars'),
		).toMatchObject({
			status: 'do-not-promote-yet',
			ownerLoops: ['performance'],
			boundary: expect.stringContaining('not a storage engine'),
		})
		expect(index.readiness).toMatchObject({
			releaseGate: 'blocked-by-publication-policy',
			headlineClaimsAllowed: false,
			implementationSurfacePromotionAllowed: false,
			implementationSurfacePromotionBoundary: expect.stringContaining(
				'do not authorize new SDK, CLI, API, or MCP surfaces',
			),
			totalRequirementCount: 9,
			missingRequirementCount: 9,
			satisfiedRequirementCount: 0,
			missingByOwnerLoop: {
				correctness: 1,
				performance: 2,
				product: 2,
				release: 4,
			},
			missingByArtifact: {
				'safe-open-proof': [
					'public-edge-fixtures',
					'release-latency-run',
					'publication-boundary',
					'compact-report-publication-policy',
				],
				'package-action-proof': [
					'edge-fixture-policy',
					'provenance-boundary',
					'unsupported-feature-boundary',
					'streaming-matrix-boundary',
					'compact-report-publication-policy',
				],
			},
		})
		expect(index.readiness.nextOwnerActions.map((action) => action.requirementId)).toEqual([
			'edge-fixture-policy',
			'public-edge-fixtures',
			'unsupported-feature-boundary',
			'release-latency-run',
			'streaming-matrix-boundary',
			'provenance-boundary',
			'publication-boundary',
			'compact-report-publication-policy',
			'compact-report-publication-policy',
		])
		expect(index.readiness.nextOwnerActions[0]).toMatchObject({
			rank: 10,
			artifact: 'package-action-proof',
			ownerLoop: 'product',
			priority: 'claim-evidence',
			nextStepKind: 'owner-decision-or-fixture-replacement',
			acceptanceEvidence: expect.stringContaining('accepts disclosed generated'),
			forbiddenShortcut: expect.stringContaining('Do not hide generated fixture provenance'),
		})
		expect(
			index.readiness.nextOwnerActions.find(
				(action) => action.requirementId === 'release-latency-run',
			),
		).toMatchObject({
			nextStepKind: 'validation-run',
			acceptanceEvidence: expect.stringContaining('tracked-clean release-environment'),
			forbiddenShortcut: expect.stringContaining('private-corpus'),
		})
		expect(
			index.readiness.nextOwnerActions.find(
				(action) => action.requirementId === 'streaming-matrix-boundary',
			),
		).toMatchObject({
			nextStepKind: 'owner-decision-or-harness-expansion',
			acceptanceEvidence: expect.stringContaining('one representative dirty-sheet streaming proof'),
			forbiddenShortcut: expect.stringContaining('full streaming parity'),
		})
		expect(index.readiness.nextOwnerActions.at(-1)).toMatchObject({
			rank: 60,
			requirementId: 'compact-report-publication-policy',
			priority: 'publication-policy',
			nextStepKind: 'publication-policy',
			acceptanceEvidence: expect.stringContaining('artifact storage path'),
			forbiddenShortcut: expect.stringContaining('Do not index or publish compact report digests'),
		})
		expect(index.readiness.implementationHandoffs).toHaveLength(2)
		const safeOpenHandoff = index.readiness.implementationHandoffs[0]
		expect(safeOpenHandoff.rank).toBe(1)
		expect(safeOpenHandoff.artifact).toBe('safe-open-proof')
		expect(safeOpenHandoff.claim).toBe('safe unknown workbook opening')
		expect(safeOpenHandoff.ownerLoops).toEqual(['performance', 'product', 'release'])
		expect(safeOpenHandoff.implementationSurfacePromotionAllowed).toBe(false)
		expect(safeOpenHandoff.proofRequired.fixture).toContain(
			'generated signed and unknown-part cases',
		)
		expect(safeOpenHandoff.proofRequired.benchmark).toContain(
			'Release-environment open-plan latency',
		)
		expect(safeOpenHandoff.proofRequired.surface).toContain('no new opener surface')
		expect(safeOpenHandoff.proofRequired.validationGate).toContain('safe-open proof harness')
		expect(safeOpenHandoff.proofRequired.competitorContrast).toContain('Microsoft Protected View')
		expect(safeOpenHandoff.proofRequired.honestBoundary).toContain('Not malware scanning')
		expect(safeOpenHandoff.proofRequired.killCriterion).toContain(
			'Do not publish headline wording if generated signed/unknown packages are hidden',
		)
		expect(safeOpenHandoff.blockingRequirementIds).toEqual([
			'public-edge-fixtures',
			'release-latency-run',
			'publication-boundary',
			'compact-report-publication-policy',
		])
		expect(safeOpenHandoff.nextStepKinds).toEqual([
			'owner-decision-or-fixture-replacement',
			'validation-run',
			'publication-policy',
		])
		expect(safeOpenHandoff.boundary).toContain(
			'not permission to add new SDK, CLI, API, or MCP surfaces',
		)
		expect(safeOpenHandoff.blockingActions.map((action) => action.requirementId)).toEqual([
			'public-edge-fixtures',
			'release-latency-run',
			'publication-boundary',
			'compact-report-publication-policy',
		])
		expect(safeOpenHandoff.blockingActions[0].acceptanceEvidence).toContain(
			'generated signed/unknown',
		)
		expect(safeOpenHandoff.blockingActions[1].forbiddenShortcut).toContain('private-corpus')
		expect(safeOpenHandoff.blockingActions[2].forbiddenShortcut).toContain('signed provenance')
		expect(safeOpenHandoff.blockingActions[3].acceptanceEvidence).toContain('artifact storage path')
		const packageActionHandoff = index.readiness.implementationHandoffs[1]
		expect(packageActionHandoff.rank).toBe(2)
		expect(packageActionHandoff.artifact).toBe('package-action-proof')
		expect(packageActionHandoff.claim).toBe('auditable package-part mutation')
		expect(packageActionHandoff.ownerLoops).toEqual([
			'correctness',
			'performance',
			'product',
			'release',
		])
		expect(packageActionHandoff.implementationSurfacePromotionAllowed).toBe(false)
		expect(packageActionHandoff.proofRequired.fixture).toContain('docProps passthrough')
		expect(packageActionHandoff.proofRequired.benchmark).toContain('Package-proof overhead')
		expect(packageActionHandoff.proofRequired.surface).toContain('no new mutation surface')
		expect(packageActionHandoff.proofRequired.validationGate).toContain(
			'package-action proof harness',
		)
		expect(packageActionHandoff.proofRequired.competitorContrast).toContain('openpyxl and SheetJS')
		expect(packageActionHandoff.proofRequired.honestBoundary).toContain('Not signed provenance')
		expect(packageActionHandoff.proofRequired.killCriterion).toContain(
			'local digests imply attestation',
		)
		expect(packageActionHandoff.blockingRequirementIds).toEqual([
			'edge-fixture-policy',
			'provenance-boundary',
			'unsupported-feature-boundary',
			'streaming-matrix-boundary',
			'compact-report-publication-policy',
		])
		expect(packageActionHandoff.nextStepKinds).toEqual([
			'owner-decision-or-fixture-replacement',
			'owner-boundary-approval',
			'owner-decision-or-harness-expansion',
			'publication-policy',
		])
		expect(packageActionHandoff.blockingActions.map((action) => action.requirementId)).toEqual([
			'edge-fixture-policy',
			'unsupported-feature-boundary',
			'streaming-matrix-boundary',
			'provenance-boundary',
			'compact-report-publication-policy',
		])
		expect(packageActionHandoff.blockingActions[0].acceptanceEvidence).toContain(
			'generated signature/unknown',
		)
		expect(packageActionHandoff.blockingActions[1].forbiddenShortcut).toContain('chart XML')
		expect(packageActionHandoff.blockingActions[2].forbiddenShortcut).toContain(
			'full streaming parity',
		)
		expect(packageActionHandoff.blockingActions[3].acceptanceEvidence).toContain('Sigstore')
		expect(packageActionHandoff.blockingActions[4].forbiddenShortcut).toContain(
			'compact report digests',
		)
		for (const artifact of index.artifacts) {
			expect(artifact.sha256).toMatch(/^[a-f0-9]{64}$/)
			expect(artifact.stableShapeSha256).toMatch(/^[a-f0-9]{64}$/)
			expect(artifact.jsonBytes).toBeGreaterThan(100)
			expect(artifact.markdownBytes).toBeGreaterThan(100)
			expect(artifact.readyWhen.length).toBeGreaterThan(0)
			expect(artifact.readyWhen.every((requirement) => requirement.status === 'missing')).toBe(true)
		}
		expect(index.artifacts.find((artifact) => artifact.name === 'safe-open-proof')).toMatchObject({
			command: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json',
			compactReportCommand:
				'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json',
			publicationStatus: 'needs-release-packaging',
			headlineClaimAllowed: false,
			releaseGate: 'blocked-by-publication-policy',
			readyWhen: [
				expect.objectContaining({
					id: 'public-edge-fixtures',
					ownerLoop: 'product',
				}),
				expect.objectContaining({
					id: 'release-latency-run',
					ownerLoop: 'performance',
				}),
				expect.objectContaining({
					id: 'publication-boundary',
					ownerLoop: 'release',
				}),
				expect.objectContaining({
					id: 'compact-report-publication-policy',
					ownerLoop: 'release',
				}),
			],
			fixtureProvenance: {
				publicFixtureCases: 6,
				generatedWorkbookCases: 0,
				generatedEdgePackageCases: 2,
				malformedCases: 1,
				generatedCaseNames: ['signed', 'unknown-part', 'malformed'],
				deterministicGeneratedCaseNames: ['signed', 'unknown-part', 'malformed'],
				generatedCaseSha256: {
					signed: expect.stringMatching(/^[a-f0-9]{64}$/),
					'unknown-part': expect.stringMatching(/^[a-f0-9]{64}$/),
					malformed: expect.stringMatching(/^[a-f0-9]{64}$/),
				},
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
			compactReportCommand:
				'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json',
			publicationStatus: 'needs-release-packaging',
			headlineClaimAllowed: false,
			releaseGate: 'blocked-by-publication-policy',
			readyWhen: [
				expect.objectContaining({
					id: 'edge-fixture-policy',
					ownerLoop: 'product',
					evidence: expect.stringContaining('fixture scan over tracked fixtures'),
				}),
				expect.objectContaining({
					id: 'provenance-boundary',
					ownerLoop: 'release',
				}),
				expect.objectContaining({
					id: 'unsupported-feature-boundary',
					ownerLoop: 'correctness',
				}),
				expect.objectContaining({
					id: 'streaming-matrix-boundary',
					ownerLoop: 'performance',
				}),
				expect.objectContaining({
					id: 'compact-report-publication-policy',
					ownerLoop: 'release',
				}),
			],
			fixtureProvenance: {
				publicFixtureCases: 4,
				generatedWorkbookCases: 2,
				generatedEdgePackageCases: 2,
				malformedCases: 0,
				generatedCaseNames: [
					'regenerate-existing-sheet',
					'add-sheet-part',
					'signature-invalidation-drop',
					'unknown-part-error',
				],
				deterministicGeneratedCaseNames: ['signature-invalidation-drop', 'unknown-part-error'],
				generatedCaseSha256: expect.objectContaining({
					'signature-invalidation-drop': expect.stringMatching(/^[a-f0-9]{64}$/),
					'unknown-part-error': expect.stringMatching(/^[a-f0-9]{64}$/),
				}),
			},
			summary: {
				cases: 8,
				allActionsCovered: true,
				sourceGraphEverywhere: true,
				streamingProofCases: 1,
				streamingRegenerateParts: 1,
			},
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

	test('emits compact owner handoff JSON without embedding full artifacts', async () => {
		const index = await runReleaseProofIndex({ includeTimings: false })
		const handoff = releaseProofOwnerHandoffIndex(index)

		expect(handoff).toMatchObject({
			releaseGate: 'blocked-by-publication-policy',
			headlineClaimsAllowed: false,
			implementationSurfacePromotionAllowed: false,
			missingRequirementCount: 9,
			boundary: expect.stringContaining('not a release artifact bundle'),
		})
		expect(handoff.fixturePolicy).toMatchObject({
			currentDecision: 'owner-approval-required',
			currentGeneratedStructuralCases: {
				'safe-open-proof': ['signed', 'unknown-part', 'malformed'],
				'package-action-proof': ['signature-invalidation-drop', 'unknown-part-error'],
			},
		})
		expect(JSON.stringify(handoff.fixturePolicy)).toContain('package-action-fixture-scan')
		expect(handoff.nextOwnerActions[0]).toMatchObject({
			requirementId: 'edge-fixture-policy',
			acceptanceEvidence: expect.stringContaining('accepts disclosed generated'),
			forbiddenShortcut: expect.stringContaining('Do not hide generated fixture provenance'),
		})
		expect(
			handoff.nextOwnerActions.find((action) => action.requirementId === 'provenance-boundary'),
		).toMatchObject({
			acceptanceEvidence: expect.stringContaining('SLSA'),
			forbiddenShortcut: expect.stringContaining('signed provenance'),
		})
		expect(handoff.implementationHandoffs.map((entry) => entry.artifact)).toEqual([
			'safe-open-proof',
			'package-action-proof',
		])
		expect(handoff.implementationHandoffs[0]).toMatchObject({
			claim: 'safe unknown workbook opening',
			proofCommand: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json',
			proofRequired: {
				surface: expect.stringContaining('no new opener surface'),
				killCriterion: expect.stringContaining('Do not publish headline wording'),
			},
			blockingActions: [
				expect.objectContaining({
					requirementId: 'public-edge-fixtures',
					forbiddenShortcut: expect.stringContaining('generated fixture provenance'),
				}),
				expect.objectContaining({
					requirementId: 'release-latency-run',
					acceptanceEvidence: expect.stringContaining('release-environment'),
				}),
				expect.objectContaining({
					requirementId: 'publication-boundary',
					acceptanceEvidence: expect.stringContaining('safe-open boundary'),
				}),
				expect.objectContaining({
					requirementId: 'compact-report-publication-policy',
				}),
			],
		})
		expect(handoff.implementationHandoffs[1].blockingActions).toHaveLength(5)
		expect(handoff.deferredClaims.map((entry) => entry.status)).toContain('do-not-promote-yet')
		expect(handoff.excludedEvidence.map((entry) => entry.name)).toEqual([
			'practical-latency-contracts',
		])
		expect(JSON.stringify(handoff)).not.toContain('"artifacts"')
		expect(JSON.stringify(handoff)).toContain('"nextOwnerActions"')
	})

	test('renders honest non-attestation boundaries', async () => {
		const markdown = releaseProofIndexMarkdown(
			await runReleaseProofIndex({ includeTimings: false }),
		)

		expect(markdown).toContain('Release Proof Evidence Index')
		expect(markdown).toContain('not signed provenance')
		expect(markdown).toContain('Publication blockers')
		expect(markdown).toContain('Compact report command')
		expect(markdown).toContain('Ready when')
		expect(markdown).toContain('Release Readiness Gate')
		expect(markdown).toContain('Implementation surface promotion allowed: false')
		expect(markdown).toContain('do not authorize new SDK, CLI, API, or MCP surfaces')
		expect(markdown).toContain('ReadyWhen requirements: total=9, missing=9, satisfied=0')
		expect(markdown).toContain(
			'Missing by owner loop: correctness=1, performance=2, product=2, release=4',
		)
		expect(markdown).toContain(
			'Missing by artifact: safe-open-proof=public-edge-fixtures,release-latency-run,publication-boundary',
		)
		expect(markdown).toContain('Next owner actions: 10:package-action-proof/edge-fixture-policy')
		expect(markdown).toContain('## Next Owner Actions')
		expect(markdown).toContain('## Fixture Policy')
		expect(markdown).toContain('Current decision: owner-approval-required')
		expect(markdown).toContain(
			'safe-open-proof=`bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json`',
		)
		expect(markdown).toContain(
			'package-action-proof=`bun run fixtures/benchmarks/package-action-fixture-scan.ts --json`',
		)
		expect(markdown).toContain('safe-open-proof=signed,unknown-part,malformed')
		expect(markdown).toContain(
			'package-action-proof=signature-invalidation-drop,unknown-part-error',
		)
		expect(markdown).toContain('edge case is package-topology-only')
		expect(markdown).toContain('Public binary fixtures are required when:')
		expect(markdown).toContain('OpenSSF Scorecard binary artifacts')
		expect(markdown).toContain(
			'| Rank | Artifact | Gate | Owner loop | Priority | Next step | Acceptance evidence | Forbidden shortcut |',
		)
		expect(markdown).toContain(
			'| 10 | package-action-proof | edge-fixture-policy | product | claim-evidence | owner-decision-or-fixture-replacement | Product accepts disclosed generated',
		)
		expect(markdown).toContain(
			'| 50 | package-action-proof | provenance-boundary | release | publication-policy | publication-policy | Release approves local package-action proof wording below SLSA',
		)
		expect(markdown).toContain('accept=Product accepts disclosed generated')
		expect(markdown).toContain('forbid=Do not hide generated fixture provenance')
		expect(markdown).toContain('tracked-clean release-environment')
		expect(markdown).toContain('Do not call local digests signed provenance')
		expect(markdown).toContain('Implementation handoffs: 1:safe-open-proof')
		expect(markdown).toContain('promotion=false;blockers=public-edge-fixtures')
		expect(markdown).toContain('kill=Do not publish headline wording')
		expect(markdown).toContain('kill=Do not publish stronger wording')
		expect(markdown).toContain('2:package-action-proof')
		expect(markdown).toContain(
			'60:safe-open-proof/compact-report-publication-policy(release,publication-policy,publication-policy;',
		)
		expect(markdown).toContain('owner-decision-or-fixture-replacement')
		expect(markdown).toContain('owner-decision-or-harness-expansion')
		expect(markdown).toContain('validation-run')
		expect(markdown).toContain('Headline claim allowed')
		expect(markdown).toContain('blocked-by-publication-policy')
		expect(markdown).toContain('public-edge-fixtures(missing,product)')
		expect(markdown).toContain('release-latency-run(missing,performance)')
		expect(markdown).toContain('edge-fixture-policy(missing,product)')
		expect(markdown).toContain('provenance-boundary(missing,release)')
		expect(markdown).toContain('streaming-matrix-boundary(missing,performance)')
		expect(markdown).toContain('compact-report-publication-policy(missing,release)')
		expect(markdown).toContain('Fixture provenance')
		expect(markdown).toContain('generatedCases=signed,unknown-part,malformed')
		expect(markdown).toContain('deterministicGenerated=signed,unknown-part,malformed')
		expect(markdown).toContain('generatedDigests=signed:')
		expect(markdown).toContain('safe-open-proof.ts --no-timings --json')
		expect(markdown).toContain('safe-open-proof.ts --no-timings --compact-json')
		expect(markdown).toContain('package-action-proof.ts --no-timings --compact-json')
		expect(markdown).toContain('SLSA')
		expect(markdown).toContain('Attestation: false')
		expect(markdown).toContain('safe unknown workbook opening')
		expect(markdown).toContain('auditable package-part mutation')
		expect(markdown).toContain('streamingProofCases=1')
		expect(markdown).toContain('Excluded Evidence')
		expect(markdown).toContain('practical-latency-contracts')
		expect(markdown).toContain('tracked-clean run')
		expect(markdown).toContain('Deferred Claims')
		expect(markdown).toContain('formula language-service primitives')
		expect(markdown).toContain('edit-producing rename is frozen')
		expect(markdown).toContain('columnar scan sidecars')
		expect(markdown).toContain('do-not-promote-yet')
	})
})
