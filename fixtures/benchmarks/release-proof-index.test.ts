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
		expect(index.fixturePolicy.approvalChecklist).toHaveLength(4)
		expect(index.fixturePolicy.approvalChecklist.map((item) => item.gateId)).toEqual([
			'public-edge-fixtures',
			'edge-fixture-policy',
			'publication-boundary',
			'provenance-boundary',
		])
		expect(index.fixturePolicy.approvalChecklist.map((item) => item.status)).toEqual([
			'pending-owner-decision',
			'pending-owner-decision',
			'pending-owner-decision',
			'pending-owner-decision',
		])
		expect(index.fixturePolicy.approvalChecklist[0]).toMatchObject({
			artifact: 'safe-open-proof',
			ownerLoop: 'product',
			validationCommand: 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
			rejectIf: expect.stringContaining('Generated fixtures are hidden'),
		})
		expect(index.fixturePolicy.approvalChecklist[3]).toMatchObject({
			artifact: 'package-action-proof',
			ownerLoop: 'release',
			validationCommand:
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
			rejectIf: expect.stringContaining('signed provenance'),
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
		expect(index.fixturePolicyEvidence).toMatchObject({
			ownerLoop: 'product',
			status: 'tracked-scan-complete-owner-approval-required',
			ownerApprovalRequired: true,
			allScansUseTrackedCorpus: true,
			publicReplacementGapsRemain: true,
			boundary: expect.stringContaining('does not prove that no suitable public fixtures exist'),
		})
		expect(index.fixturePolicyEvidence.safeOpen).toMatchObject({
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			validationCommand: 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
			corpus: 'tracked-git-fixtures',
			replacementStatus: 'no-public-binary-replacement-found',
			signatureOrUnknownMatches: 0,
			currentGeneratedStructuralCases: ['signed', 'unknown-part', 'malformed'],
		})
		expect(index.fixturePolicyEvidence.safeOpen.scanned).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.packageAction).toMatchObject({
			artifact: 'package-action-proof',
			gateId: 'edge-fixture-policy',
			validationCommand: 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
			corpus: 'tracked-git-fixtures',
			replacementStatus: 'remaining-generated-edge-cases',
			currentGeneratedStructuralCases: ['signature-invalidation-drop', 'unknown-part-error'],
			missingReplacementFeatures: ['signaturePackage', 'syntheticUnknownPathFamily'],
		})
		expect(index.fixturePolicyEvidence.packageAction.scanned).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.packageAction.featureCounts).toMatchObject({
			signaturePackage: 0,
			syntheticUnknownPathFamily: 0,
		})
		expect(index.fixturePolicyEvidence.packageAction.featureCounts.docPropsCore).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.packageAction.featureCounts.calcChain).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.packageAction.featureCounts.macro).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.packageAction.featureCounts.chartOrDrawing).toBeGreaterThan(
			0,
		)
		expect(index.generatedFixtureDecisionEvidence).toMatchObject({
			ownerLoop: 'product',
			status: 'generated-structural-cases-disclosed-owner-approval-required',
			ownerApprovalRequired: true,
			allGeneratedStructuralCasesDisclosed: true,
			publicReplacementGapsRemain: true,
			validationCommand:
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
			boundary: expect.stringContaining('does not approve generated fixtures'),
		})
		expect(index.generatedFixtureDecisionEvidence.cases.map((entry) => entry.caseName)).toEqual([
			'signed',
			'unknown-part',
			'malformed',
			'signature-invalidation-drop',
			'unknown-part-error',
		])
		expect(index.generatedFixtureDecisionEvidence.cases[0]).toMatchObject({
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			generatedKind: 'generated-edge-package',
			replacementEvidence: expect.stringContaining('signatureOrUnknownMatches=0'),
			forbiddenUse: expect.stringContaining('signature verification'),
		})
		expect(index.generatedFixtureDecisionEvidence.cases[2]).toMatchObject({
			caseName: 'malformed',
			generatedKind: 'generated-malformed-package',
			allowedUse: expect.stringContaining('malformed-package rejection'),
			forbiddenUse: expect.stringContaining('arbitrary malformed files'),
		})
		expect(index.generatedFixtureDecisionEvidence.cases[3]).toMatchObject({
			artifact: 'package-action-proof',
			gateId: 'edge-fixture-policy',
			replacementEvidence: expect.stringContaining('signaturePackage=0'),
			forbiddenUse: expect.stringContaining('SLSA'),
		})
		expect(index.generatedFixtureDecisionEvidence.cases[4]).toMatchObject({
			replacementEvidence: expect.stringContaining('syntheticUnknownPathFamily=0'),
			allowedUse: expect.stringContaining('explicit error action'),
		})
		expect(index.performancePolicy).toMatchObject({
			currentDecision: 'owner-approval-required',
			boundary: expect.stringContaining('not a release performance threshold'),
		})
		expect(index.performancePolicy.approvalChecklist).toHaveLength(2)
		expect(index.performancePolicy.approvalChecklist.map((item) => item.gateId)).toEqual([
			'release-latency-run',
			'streaming-matrix-boundary',
		])
		expect(index.performancePolicy.approvalChecklist[0]).toMatchObject({
			artifact: 'safe-open-proof',
			ownerLoop: 'performance',
			status: 'pending-owner-decision',
			validationCommand:
				'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 3 --warmup 1 --json',
			rejectIf: expect.stringContaining('latency SLA'),
		})
		expect(index.performancePolicy.approvalChecklist[1]).toMatchObject({
			artifact: 'package-action-proof',
			ownerLoop: 'performance',
			status: 'pending-owner-decision',
			validationCommand: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
			rejectIf: expect.stringContaining('full streaming parity'),
		})
		expect(index.performancePolicy.sourceReferences.map((entry) => entry.label)).toEqual([
			'Bun benchmarking',
			'hyperfine benchmarking',
			'hyperfine manual',
		])
		expect(index.streamingMatrixEvidence).toMatchObject({
			artifact: 'package-action-proof',
			gateId: 'streaming-matrix-boundary',
			ownerLoop: 'performance',
			status: 'representative-proof-present-owner-approval-required',
			ownerApprovalRequired: true,
			validationCommand: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
			representativeProofCases: 1,
			streamingRegenerateParts: 1,
			coveredActionKinds: ['passthrough', 'regenerate'],
			missingActionKinds: ['add', 'drop', 'error'],
			coveredCaseNames: ['docprops-passthrough'],
			streamingIssueCaseNames: [],
			boundary: expect.stringContaining('does not prove full streaming parity'),
		})
		expect(index.streamingMatrixEvidence.nonStreamingCaseNames).toEqual([
			'regenerate-existing-sheet',
			'add-sheet-part',
			'calc-chain-drop',
			'signature-invalidation-drop',
			'macro-passthrough',
			'chart-sidecar-accounting',
			'unknown-part-error',
		])
		expect(index.streamingMatrixEvidence.publicNonStreamingCaseNames).toEqual([
			'calc-chain-drop',
			'macro-passthrough',
			'chart-sidecar-accounting',
		])
		expect(index.streamingMatrixEvidence.generatedNonStreamingCaseNames).toEqual([
			'regenerate-existing-sheet',
			'add-sheet-part',
			'signature-invalidation-drop',
			'unknown-part-error',
		])
		expect(index.correctnessPolicy).toMatchObject({
			currentDecision: 'owner-approval-required',
			boundary: expect.stringContaining('does not prove semantic support'),
		})
		expect(index.correctnessPolicy.approvalChecklist).toHaveLength(1)
		expect(index.correctnessPolicy.approvalChecklist[0]).toMatchObject({
			artifact: 'package-action-proof',
			gateId: 'unsupported-feature-boundary',
			ownerLoop: 'correctness',
			status: 'pending-owner-decision',
			validationCommand: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
			rejectIf: expect.stringContaining('signature preservation or verification'),
		})
		expect(
			index.correctnessPolicy.unsupportedFeatureBoundaries.map((item) => item.feature),
		).toEqual([
			'digital-signatures',
			'calc-chain',
			'chart-drawing-sidecars',
			'macros-activex',
			'unknown-parts',
			'streaming-scope',
		])
		expect(index.correctnessPolicy.unsupportedFeatureBoundaries[0]).toMatchObject({
			allowedWording: expect.stringContaining('detects signature package parts'),
			forbiddenWording: expect.stringContaining('verifies'),
		})
		expect(index.correctnessPolicy.unsupportedFeatureBoundaries[2]).toMatchObject({
			allowedWording: expect.stringContaining('accounts for chart/drawing sidecars'),
			forbiddenWording: expect.stringContaining('byte-passthrough'),
		})
		expect(index.correctnessPolicy.unsupportedFeatureBoundaries[5]).toMatchObject({
			allowedWording: expect.stringContaining('representative streaming writer'),
			forbiddenWording: expect.stringContaining('Full streaming parity'),
		})
		expect(index.correctnessPolicy.sourceReferences.map((entry) => entry.label)).toEqual([
			'OOXML calculation chain',
			'Microsoft macro security',
			'Microsoft ActiveX settings',
			'SheetJS VBA blobs',
			'OOXML digital signatures',
		])
		expect(index.correctnessBoundaryEvidence).toMatchObject({
			artifact: 'package-action-proof',
			gateId: 'unsupported-feature-boundary',
			ownerLoop: 'correctness',
			status: 'evidence-present-owner-approval-required',
			allCurrentEvidencePresent: true,
			missingFeatureNames: [],
			ownerEscalationRequired: false,
			ownerApprovalRequired: true,
			validationCommand:
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
			boundary: expect.stringContaining('does not satisfy the owner gate'),
		})
		expect(index.correctnessBoundaryEvidence.featureChecks.map((item) => item.feature)).toEqual([
			'digital-signatures',
			'calc-chain',
			'chart-drawing-sidecars',
			'macros-activex',
			'unknown-parts',
			'streaming-scope',
		])
		expect(
			index.correctnessBoundaryEvidence.featureChecks.every((item) => item.evidencePresent),
		).toBe(true)
		expect(index.correctnessBoundaryEvidence.featureChecks[0]).toMatchObject({
			evidenceSources: [
				'package-action-proof/signature-invalidation-drop',
				'safe-open-proof/signed',
			],
			proofChecks: expect.arrayContaining([
				expect.stringContaining('drop action for signature package parts'),
			]),
			forbiddenWording: expect.stringContaining('re-signs'),
		})
		expect(index.correctnessBoundaryEvidence.featureChecks[3]).toMatchObject({
			evidenceSources: [
				'package-action-proof/macro-passthrough',
				'safe-open-proof/macro',
				'safe-open-proof/activex',
			],
			proofChecks: expect.arrayContaining([
				expect.stringContaining('macro and ActiveX risk families'),
			]),
			forbiddenWording: expect.stringContaining('safe, sandboxed'),
		})
		expect(index.correctnessBoundaryEvidence.featureChecks[4]).toMatchObject({
			proofChecks: expect.arrayContaining([expect.stringContaining('fails closed')]),
			allowedWording: expect.stringContaining('explicit unknown-part error'),
		})
		expect(index.compactReportPublicationEvidence).toMatchObject({
			ownerLoop: 'release',
			status: 'local-summary-present-publication-policy-required',
			ownerApprovalRequired: true,
			compactReportDigestsIndexed: false,
			allCompactCommandsPresent: true,
			compactReportsEmbedForbiddenPayloadFields: false,
			generatedAtIncluded: true,
			missingPolicyRequirements: [
				'artifact storage path',
				'retention and privacy filtering',
				'canonicalization subject',
				'offline verification expectations',
			],
			boundary: expect.stringContaining('does not publish compact report digests'),
		})
		expect(index.compactReportPublicationEvidence.policyDecisions).toHaveLength(4)
		expect(
			index.compactReportPublicationEvidence.policyDecisions.map(
				(decision) => decision.requirement,
			),
		).toEqual(index.compactReportPublicationEvidence.missingPolicyRequirements)
		expect(
			index.compactReportPublicationEvidence.policyDecisions.map((item) => item.status),
		).toEqual([
			'pending-owner-decision',
			'pending-owner-decision',
			'pending-owner-decision',
			'pending-owner-decision',
		])
		expect(index.compactReportPublicationEvidence.policyDecisions[0]).toMatchObject({
			requirement: 'artifact storage path',
			acceptanceEvidence: expect.stringContaining('path convention'),
			rejectIf: expect.stringContaining('temporary paths'),
		})
		expect(index.compactReportPublicationEvidence.policyDecisions[1]).toMatchObject({
			requirement: 'retention and privacy filtering',
			rejectIf: expect.stringContaining('workbook bytes'),
		})
		expect(index.compactReportPublicationEvidence.policyDecisions[2]).toMatchObject({
			requirement: 'canonicalization subject',
			acceptanceEvidence: expect.stringContaining('canonical JSON subject'),
		})
		expect(index.compactReportPublicationEvidence.policyDecisions[3]).toMatchObject({
			requirement: 'offline verification expectations',
			rejectIf: expect.stringContaining('signed provenance'),
		})
		expect(index.compactReportPublicationEvidence.reports.map((item) => item.artifact)).toEqual([
			'safe-open-proof',
			'package-action-proof',
		])
		expect(index.compactReportPublicationEvidence.reports[0]).toMatchObject({
			gateId: 'compact-report-publication-policy',
			command: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json',
			forbiddenPayloadFieldsPresent: [],
			readyWhenGatePresent: true,
			generatedAtIncluded: true,
			headlineClaimAllowed: false,
			releaseGate: 'blocked-by-publication-policy',
		})
		expect(index.compactReportPublicationEvidence.reports[1]).toMatchObject({
			gateId: 'compact-report-publication-policy',
			command: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json',
			forbiddenPayloadFieldsPresent: [],
			readyWhenGatePresent: true,
			generatedAtIncluded: true,
			headlineClaimAllowed: false,
			releaseGate: 'blocked-by-publication-policy',
		})
		expect(index.compactReportPublicationEvidence.reports[0].topLevelFields).toContain('cases')
		expect(index.compactReportPublicationEvidence.reports[1].topLevelFields).toContain('coverage')
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
		expect(handoff.fixturePolicy.approvalChecklist.map((item) => item.ownerLoop)).toEqual([
			'product',
			'product',
			'release',
			'release',
		])
		expect(handoff.fixturePolicyEvidence).toMatchObject({
			status: 'tracked-scan-complete-owner-approval-required',
			allScansUseTrackedCorpus: true,
			publicReplacementGapsRemain: true,
			ownerApprovalRequired: true,
		})
		expect(handoff.fixturePolicyEvidence.safeOpen.signatureOrUnknownMatches).toBe(0)
		expect(handoff.fixturePolicyEvidence.packageAction.missingReplacementFeatures).toEqual([
			'signaturePackage',
			'syntheticUnknownPathFamily',
		])
		expect(handoff.generatedFixtureDecisionEvidence).toMatchObject({
			status: 'generated-structural-cases-disclosed-owner-approval-required',
			allGeneratedStructuralCasesDisclosed: true,
			publicReplacementGapsRemain: true,
			ownerApprovalRequired: true,
		})
		expect(handoff.generatedFixtureDecisionEvidence.cases.map((entry) => entry.caseName)).toEqual([
			'signed',
			'unknown-part',
			'malformed',
			'signature-invalidation-drop',
			'unknown-part-error',
		])
		expect(handoff.performancePolicy.approvalChecklist.map((item) => item.gateId)).toEqual([
			'release-latency-run',
			'streaming-matrix-boundary',
		])
		expect(handoff.streamingMatrixEvidence).toMatchObject({
			status: 'representative-proof-present-owner-approval-required',
			ownerApprovalRequired: true,
			coveredActionKinds: ['passthrough', 'regenerate'],
			missingActionKinds: ['add', 'drop', 'error'],
			coveredCaseNames: ['docprops-passthrough'],
			publicNonStreamingCaseNames: [
				'calc-chain-drop',
				'macro-passthrough',
				'chart-sidecar-accounting',
			],
		})
		expect(handoff.correctnessPolicy.approvalChecklist.map((item) => item.gateId)).toEqual([
			'unsupported-feature-boundary',
		])
		expect(
			handoff.correctnessPolicy.unsupportedFeatureBoundaries.map((item) => item.feature),
		).toEqual([
			'digital-signatures',
			'calc-chain',
			'chart-drawing-sidecars',
			'macros-activex',
			'unknown-parts',
			'streaming-scope',
		])
		expect(handoff.correctnessBoundaryEvidence).toMatchObject({
			status: 'evidence-present-owner-approval-required',
			allCurrentEvidencePresent: true,
			missingFeatureNames: [],
			ownerEscalationRequired: false,
			ownerApprovalRequired: true,
		})
		expect(handoff.correctnessBoundaryEvidence.featureChecks.map((item) => item.feature)).toEqual([
			'digital-signatures',
			'calc-chain',
			'chart-drawing-sidecars',
			'macros-activex',
			'unknown-parts',
			'streaming-scope',
		])
		expect(JSON.stringify(handoff.fixturePolicy)).toContain('package-action-fixture-scan')
		expect(JSON.stringify(handoff.generatedFixtureDecisionEvidence)).toContain(
			'generated-malformed-package',
		)
		expect(JSON.stringify(handoff.performancePolicy)).toContain('safe-open-proof.ts --repeat 3')
		expect(JSON.stringify(handoff.streamingMatrixEvidence)).toContain(
			'"missingActionKinds":["add","drop","error"]',
		)
		expect(JSON.stringify(handoff.correctnessPolicy)).toContain('signature preservation')
		expect(JSON.stringify(handoff.correctnessBoundaryEvidence)).toContain('safe-open-proof/activex')
		expect(handoff.compactReportPublicationEvidence).toMatchObject({
			status: 'local-summary-present-publication-policy-required',
			compactReportDigestsIndexed: false,
			allCompactCommandsPresent: true,
			compactReportsEmbedForbiddenPayloadFields: false,
			ownerApprovalRequired: true,
		})
		expect(handoff.compactReportPublicationEvidence.reports.map((item) => item.artifact)).toEqual([
			'safe-open-proof',
			'package-action-proof',
		])
		expect(
			handoff.compactReportPublicationEvidence.policyDecisions.map(
				(decision) => decision.requirement,
			),
		).toEqual([
			'artifact storage path',
			'retention and privacy filtering',
			'canonicalization subject',
			'offline verification expectations',
		])
		expect(JSON.stringify(handoff.compactReportPublicationEvidence)).toContain(
			'offline verification expectations',
		)
		expect(JSON.stringify(handoff.compactReportPublicationEvidence)).toContain(
			'canonical JSON subject',
		)
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
		expect(markdown).toContain('Approval checklist:')
		expect(markdown).toContain(
			'| Artifact | Gate | Owner | Status | Decision needed | Acceptance evidence | Reject if | Validation command |',
		)
		expect(markdown).toContain('| safe-open-proof | public-edge-fixtures | product')
		expect(markdown).toContain('| package-action-proof | provenance-boundary | release')
		expect(markdown).toContain('pending-owner-decision')
		expect(markdown).toContain('OpenSSF Scorecard binary artifacts')
		expect(markdown).toContain('Fixture policy evidence:')
		expect(markdown).toContain('Status: tracked-scan-complete-owner-approval-required')
		expect(markdown).toContain('All scans use tracked corpus: true')
		expect(markdown).toContain('Public replacement gaps remain: true')
		expect(markdown).toContain('| safe-open-proof | public-edge-fixtures')
		expect(markdown).toContain('signatureOrUnknownMatches=0')
		expect(markdown).toContain('| package-action-proof | edge-fixture-policy')
		expect(markdown).toContain(
			'missingReplacementFeatures=signaturePackage,syntheticUnknownPathFamily',
		)
		expect(markdown).toContain('Generated fixture decision evidence:')
		expect(markdown).toContain(
			'Status: generated-structural-cases-disclosed-owner-approval-required',
		)
		expect(markdown).toContain('All generated structural cases disclosed: true')
		expect(markdown).toContain('| safe-open-proof | public-edge-fixtures | signed')
		expect(markdown).toContain('| safe-open-proof | public-edge-fixtures | malformed')
		expect(markdown).toContain(
			'| package-action-proof | edge-fixture-policy | signature-invalidation-drop',
		)
		expect(markdown).toContain('tracked package-action scan found signaturePackage=0')
		expect(markdown).toContain('Do not claim recovery of arbitrary malformed files')
		expect(markdown).toContain('## Performance Policy')
		expect(markdown).toContain('not a release performance threshold')
		expect(markdown).toContain('| safe-open-proof | release-latency-run | performance')
		expect(markdown).toContain('| package-action-proof | streaming-matrix-boundary | performance')
		expect(markdown).toContain('Bun benchmarking')
		expect(markdown).toContain('## Streaming Matrix Evidence')
		expect(markdown).toContain('Status: representative-proof-present-owner-approval-required')
		expect(markdown).toContain('Covered action kinds: passthrough,regenerate')
		expect(markdown).toContain('Missing action kinds: add,drop,error')
		expect(markdown).toContain('Covered cases: docprops-passthrough')
		expect(markdown).toContain('Non-streaming cases: regenerate-existing-sheet,add-sheet-part')
		expect(markdown).toContain(
			'Public non-streaming cases: calc-chain-drop,macro-passthrough,chart-sidecar-accounting',
		)
		expect(markdown).toContain('## Correctness Policy')
		expect(markdown).toContain('Unsupported feature boundaries:')
		expect(markdown).toContain(
			'| Feature | Current evidence | Allowed wording | Forbidden wording |',
		)
		expect(markdown).toContain('| digital-signatures | Generated signature-invalidation')
		expect(markdown).toContain(
			'| package-action-proof | unsupported-feature-boundary | correctness',
		)
		expect(markdown).toContain('signature preservation or verification')
		expect(markdown).toContain('Chart XML is byte-passthrough')
		expect(markdown).toContain('OOXML digital signatures')
		expect(markdown).toContain('Correctness boundary evidence:')
		expect(markdown).toContain('Status: evidence-present-owner-approval-required')
		expect(markdown).toContain('All current evidence present: true')
		expect(markdown).toContain('Missing feature names: none')
		expect(markdown).toContain('Owner escalation required: false')
		expect(markdown).toContain('Owner approval required: true')
		expect(markdown).toContain('does not satisfy the owner gate')
		expect(markdown).toContain(
			'| Feature | Evidence present | Sources | Checks | Allowed wording | Forbidden wording |',
		)
		expect(markdown).toContain('| macros-activex | true | package-action-proof/macro-passthrough')
		expect(markdown).toContain('safe-open-proof/activex')
		expect(markdown).toContain('post-write audit fails closed')
		expect(markdown).toContain('## Compact Report Publication Evidence')
		expect(markdown).toContain('Status: local-summary-present-publication-policy-required')
		expect(markdown).toContain('Compact report digests indexed: false')
		expect(markdown).toContain('Forbidden payload fields embedded: false')
		expect(markdown).toContain('GeneratedAt included: true')
		expect(markdown).toContain('Missing publication policy requirements:')
		expect(markdown).toContain('offline verification expectations')
		expect(markdown).toContain('Publication policy decisions:')
		expect(markdown).toContain(
			'| Requirement | Owner | Status | Decision needed | Acceptance evidence | Reject if |',
		)
		expect(markdown).toContain('| artifact storage path | release | pending-owner-decision')
		expect(markdown).toContain('| canonicalization subject | release | pending-owner-decision')
		expect(markdown).toContain('canonical JSON subject')
		expect(markdown).toContain('transparency-log inclusion')
		expect(markdown).toContain('| safe-open-proof | compact-report-publication-policy')
		expect(markdown).toContain('| package-action-proof | compact-report-publication-policy')
		expect(markdown).toContain('does not publish compact report digests')
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
