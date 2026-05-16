import { describe, expect, test } from 'bun:test'
import { fileURLToPath } from 'node:url'
import {
	releaseProofIndexMarkdown,
	releaseProofOwnerHandoffIndex,
	runReleaseProofIndex,
} from './release-proof-index.ts'

const runnerPath = fileURLToPath(new URL('./release-proof-index.ts', import.meta.url))

describe('release proof evidence index', () => {
	test('references top claim proof artifacts by digest without embedding artifacts', async () => {
		const index = await runReleaseProofIndex({ includeTimings: false })

		expect(index.signed).toBe(false)
		expect(index.attestation).toBe(false)
		expect(index.releasePackageabilityEvidence).toMatchObject({
			ownerLoop: 'release',
			status: 'local-tarball-smokes-present-publication-policy-required',
			ownerApprovalRequired: true,
			sdkSmokeCommand: 'bun run release:sdk:smoke',
			appSmokeCommand: 'bun run release:apps:smoke',
			rcGateCommand: 'bun run release:rc:gate',
			boundary: expect.stringContaining('local tarball install'),
		})
		expect(index.releasePackageabilityEvidence.coveredEvidence).toEqual([
			expect.stringContaining('SDK tarball installs'),
			expect.stringContaining('CLI/API/MCP app tarballs install'),
			expect.stringContaining('Unified RC gate builds JS artifacts'),
			expect.stringContaining('Installed CLI bin reports version, completes'),
			expect.stringContaining(
				'Installed API createApiFetch handles write/inspect/plan/commit/check/read',
			),
			expect.stringContaining('Installed MCP package registers tool/resource callbacks'),
		])
		expect(index.releasePackageabilityEvidence.missingPolicyRequirements).toContain(
			'registry publication workflow',
		)
		expect(index.releasePackageabilityEvidence.forbiddenClaims).toContain('signed provenance')
		expect(index.excludedEvidenceCount).toBe(1)
		expect(index.deferredClaimCount).toBe(7)
		expect(index.fixturePolicy).toMatchObject({
			currentDecision: 'owner-approval-required',
			trackedFixtureScanCommands: {
				'safe-open-proof': 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
				'package-action-proof': 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
			},
			currentGeneratedStructuralCases: {
				'safe-open-proof': ['signed', 'malformed'],
				'package-action-proof': ['signature-invalidation-drop'],
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
		expect(index.fixturePolicy.safeOpenFixtureAcceptanceChecklist).toEqual([
			expect.objectContaining({
				caseName: 'signed',
				generatedCaseKind: 'generated-edge-package',
				gateEffect: 'keeps-public-edge-fixtures-missing-until-owner-approval',
				requiresPublicBinaryWhen: expect.stringContaining('real signed workbook behavior'),
			}),
			expect.objectContaining({
				caseName: 'malformed',
				generatedCaseKind: 'generated-malformed-package',
				requiresPublicBinaryWhen: expect.stringContaining('vendor repair equivalence'),
			}),
		])
		expect(index.fixturePolicy.packageActionFixtureAcceptanceChecklist).toEqual([
			expect.objectContaining({
				caseName: 'signature-invalidation-drop',
				generatedCaseKind: 'generated-edge-package',
				gateEffect: 'keeps-edge-fixture-policy-missing-until-owner-approval',
				forbiddenClaim: expect.stringContaining('signed provenance'),
			}),
		])
		expect(index.fixturePolicy.sourceReferences.map((entry) => entry.label)).toEqual([
			'GitHub repository limits',
			'GitHub large files',
			'OpenSSF Scorecard binary artifacts',
			'SLSA 1.2 build provenance distribution',
			'GitHub artifact attestations',
		])
		expect(index.fixturePolicy.sourceReferences.map((entry) => entry.url)).toContain(
			'https://slsa.dev/spec/v1.2/distributing-provenance',
		)
		expect(index.fixturePolicy.sourceReferences.map((entry) => entry.url).join('\n')).not.toContain(
			'v1.0',
		)
		expect(index.fixturePolicyEvidence).toMatchObject({
			ownerLoop: 'product',
			status: 'tracked-scan-complete-owner-approval-required',
			ownerApprovalRequired: true,
			allScansUseTrackedCorpus: true,
			publicReplacementGapsRemain: true,
			boundary: expect.stringContaining('does not prove that no suitable public fixtures exist'),
		})
		const safeOpenRiskFamilyCounts = { ...index.fixturePolicyEvidence.safeOpen.riskFamilyCounts }
		expect(index.fixturePolicyEvidence.safeOpen).toMatchObject({
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			validationCommand: 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
			corpus: 'tracked-git-fixtures',
			replacementStatus: 'candidate-found',
			riskFamilyCounts: expect.objectContaining({
				preservedMacro: expect.any(Number),
				preservedActiveX: expect.any(Number),
			}),
			signatureOrUnknownMatches: 1,
			currentGeneratedStructuralCases: ['signed', 'malformed'],
		})
		expect(index.fixturePolicyEvidence.safeOpen.scanned).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.safeOpen.rejectedFixtures).toEqual([
			'fixtures/xlsx/calamine/pass_protected.xlsx',
		])
		expect(
			Object.values(safeOpenRiskFamilyCounts).reduce((sum, count) => sum + count, 0),
		).toBeGreaterThan(0)
		expect(safeOpenRiskFamilyCounts.preservedSignature ?? 0).toBe(0)
		expect(index.fixturePolicyEvidence.packageAction).toMatchObject({
			artifact: 'package-action-proof',
			gateId: 'edge-fixture-policy',
			validationCommand: 'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
			corpus: 'tracked-git-fixtures',
			replacementStatus: 'remaining-generated-edge-cases',
			currentGeneratedStructuralCases: ['signature-invalidation-drop'],
			missingReplacementFeatures: ['signaturePackage'],
		})
		expect(index.fixturePolicyEvidence.packageAction.scanned).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.packageAction.featureCounts).toMatchObject({
			signaturePackage: 0,
			unknownPathFamily: expect.any(Number),
		})
		expect(index.fixturePolicyEvidence.packageAction.missingReplacementFeatures).not.toContain(
			'unknownPathFamily',
		)
		expect(index.fixturePolicyEvidence.packageAction.featureCounts.docPropsCore).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.packageAction.featureCounts.calcChain).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.packageAction.featureCounts.macro).toBeGreaterThan(0)
		expect(index.fixturePolicyEvidence.packageAction.featureCounts.chartOrDrawing).toBeGreaterThan(
			0,
		)
		expect(index.fixtureAcquisitionPlan).toMatchObject({
			ownerLoop: 'product',
			status: 'ranked-owner-review-required',
			taskCount: 2,
			boundary: expect.stringContaining('not fixture approval'),
		})
		expect(index.fixtureAcquisitionPlan.tasks.map((task) => task.caseName)).toEqual([
			'signed-package',
			'malformed-package',
		])
		expect(index.fixtureAcquisitionPlan.tasks[0]).toMatchObject({
			rank: 1,
			relatedArtifacts: ['safe-open-proof', 'package-action-proof'],
			relatedGates: ['public-edge-fixtures', 'edge-fixture-policy'],
			caseName: 'signed-package',
			evidenceAlreadyPresent: expect.stringContaining('signaturePackage=0'),
			boundary: expect.stringContaining('signature validation'),
		})
		const malformedFixtureEvidence = index.fixtureAcquisitionPlan.tasks[1].evidenceAlreadyPresent
		expect(index.fixtureAcquisitionPlan.tasks[1]).toMatchObject({
			caseName: 'malformed-package',
			evidenceAlreadyPresent: expect.stringContaining('pass_protected.xlsx'),
		})
		expect(malformedFixtureEvidence).toEqual(
			expect.stringContaining('not a public malformed-package replacement'),
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
			'malformed',
			'signature-invalidation-drop',
		])
		expect(index.generatedFixtureDecisionEvidence.cases[0]).toMatchObject({
			artifact: 'safe-open-proof',
			gateId: 'public-edge-fixtures',
			generatedKind: 'generated-edge-package',
			replacementEvidence: expect.stringContaining('signatureOrUnknownMatches=1'),
			recommendedOwnerAction: expect.stringContaining('local safe-open package-feature routing'),
			forbiddenUse: expect.stringContaining('signature verification'),
		})
		expect(index.generatedFixtureDecisionEvidence.cases[1]).toMatchObject({
			caseName: 'malformed',
			generatedKind: 'generated-malformed-package',
			replacementEvidence: expect.stringContaining('pass_protected.xlsx'),
			recommendedOwnerAction: expect.stringContaining('fail-closed rejection-path proof'),
			allowedUse: expect.stringContaining('malformed-package rejection'),
			forbiddenUse: expect.stringContaining('arbitrary malformed files'),
		})
		expect(index.generatedFixtureDecisionEvidence.cases[2]).toMatchObject({
			artifact: 'package-action-proof',
			gateId: 'edge-fixture-policy',
			replacementEvidence: expect.stringContaining('signaturePackage=0'),
			recommendedOwnerAction: expect.stringContaining('local signature-part drop'),
			forbiddenUse: expect.stringContaining('SLSA'),
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
				'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
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
			'Google Benchmark repeated statistics',
		])
		expect(index.safeOpenLatencyValidationEvidence).toMatchObject({
			artifact: 'safe-open-proof',
			gateId: 'release-latency-run',
			ownerLoop: 'performance',
			status: 'timed-evidence-absent-owner-run-required',
			ownerApprovalRequired: true,
			releaseClaimAllowed: false,
			thresholdClaimAllowed: false,
			validationCommand:
				'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
			runProfile: expect.objectContaining({
				profileId: 'safe-open-release-latency-owner-review',
				command: 'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
				minimumRepeat: 10,
				minimumWarmup: 3,
				requiredPublicCaseNames: ['clean', 'formula-heavy', 'macro', 'pivot', 'activex', 'chart'],
				requireTimingEnvironment: true,
				cvGuard: expect.objectContaining({
					metric: 'publicOpenPlanCv',
					maxRecommendedCv: 0.25,
				}),
				forbiddenUses: expect.arrayContaining(['release threshold', 'QSS performance comparison']),
			}),
			runProfileSatisfied: false,
			runProfileFailures: expect.arrayContaining([
				'repeat 1 below profile minimum 10',
				'warmup 2 below profile minimum 3',
				'timing environment metadata missing',
				'required public case clean missing timed evidence',
			]),
			repeat: 1,
			warmup: 2,
			timingEnvironmentCaptured: false,
			timedCaseCount: 0,
			publicTimedCaseNames: [],
			generatedTimedCaseNames: [],
			publicOpenPlanMedianMs: {},
			publicOpenPlanP95Ms: {},
			publicOpenPlanCv: {},
			publicFullOpenMedianMs: {},
			publicFullOpenP95Ms: {},
			publicFullOpenCv: {},
			publicFullOpenRatio: {},
			malformedRejected: true,
			missingPolicyRequirements: [
				'tracked-clean release environment',
				'standardized public input set',
				'performance-owner approval of the safe-open release-latency owner-review profile',
				'non-threshold release wording',
			],
			boundary: expect.stringContaining('not a release threshold'),
		})
		expect(index.streamingMatrixEvidence).toMatchObject({
			artifact: 'package-action-proof',
			gateId: 'streaming-matrix-boundary',
			ownerLoop: 'performance',
			status: 'representative-proof-present-owner-approval-required',
			ownerApprovalRequired: true,
			validationCommand: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
			representativeProofCases: 5,
			streamingRegenerateParts: 4,
			coveredActionKinds: ['passthrough', 'regenerate', 'add', 'drop'],
			missingActionKinds: ['error'],
			coveredCaseNames: [
				'docprops-passthrough',
				'add-sheet-part',
				'calc-chain-drop',
				'macro-passthrough',
				'chart-sidecar-accounting',
			],
			streamingIssueCaseNames: [],
			boundary: expect.stringContaining('does not prove full streaming parity'),
		})
		expect(index.streamingMatrixEvidence.nonStreamingCaseNames).toEqual([
			'regenerate-existing-sheet',
			'signature-invalidation-drop',
			'unknown-part-error',
		])
		expect(index.streamingMatrixEvidence.publicNonStreamingCaseNames).toEqual([
			'unknown-part-error',
		])
		expect(index.streamingMatrixEvidence.generatedNonStreamingCaseNames).toEqual([
			'regenerate-existing-sheet',
			'signature-invalidation-drop',
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
			'encrypted-files',
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
			allowedWording: expect.stringContaining('decrypt supported OOXML password-protected'),
			forbiddenWording: expect.stringContaining('recovers passwords'),
		})
		expect(index.correctnessPolicy.unsupportedFeatureBoundaries[6]).toMatchObject({
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
			'encrypted-files',
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
		expect(index.trustCompletenessBoundaryEvidence).toMatchObject({
			ownerLoop: 'correctness',
			status: 'boundary-pinned-owner-scope',
			validationCommand: 'bun test packages/sdk/src/release-trust-matrix.test.ts',
			releaseTrustMatrixPath: 'docs/RELEASE_TRUST_MATRIX.md',
			requiredPromotionEvidence: expect.stringContaining('top release claim'),
			doesNotCloseGates: ['product', 'performance', 'release'],
			boundary: expect.stringContaining('does not approve headline release claims'),
		})
		expect(
			index.trustCompletenessBoundaryEvidence.outOfScopeClasses.map((item) => item.name),
		).toEqual([
			'Broad formula function coverage',
			'Product/DX orchestration such as progressive open or viewport merge helpers',
			'Reader/writer performance and benchmark tuning',
			'More malformed-field enumeration',
			'New unknown Excel feature implementation',
		])
		expect(
			index.trustCompletenessBoundaryEvidence.sourceReferences.map((entry) => entry.label),
		).toEqual([
			'SLSA 1.2 build provenance distribution',
			'GitHub artifact attestations',
			'Microsoft Protected View',
		])
		expect(
			index.trustCompletenessBoundaryEvidence.sourceReferences.map((entry) => entry.url),
		).toContain('https://slsa.dev/spec/v1.2/distributing-provenance')
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
			'research-surface-hygiene',
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
		expect(index.claimPortfolio.map((claim) => `${claim.rank}:${claim.name}`)).toEqual([
			'1:safe-open-proof',
			'2:package-action-proof',
			'3:formula-language-service-primitives',
			'4:token-bounded-agent-view',
			'5:retained-viewport-patch-history',
			'6:release-proof-bundle',
			'7:formula-oracle-routing',
			'8:property-journal-laws',
			'9:columnar-scan-sidecars',
			'10:agent-workflow-observability',
			'11:research-surface-hygiene',
		])
		expect(index.claimPortfolio[0]).toMatchObject({
			claim: 'safe unknown workbook opening',
			status: 'claim-wording-allowed-today',
			handoffDecision: 'top-implementation-handoff',
			likelyHandoffOwner: ['product', 'performance', 'release'],
			evidenceNeeded: {
				fixture: expect.stringContaining('Public clean'),
				killCriterion: expect.stringContaining('Do not publish headline wording'),
			},
		})
		expect(index.claimPortfolio[2]).toMatchObject({
			name: 'formula-language-service-primitives',
			status: 'needs-one-more-fold-in',
			handoffDecision: 'proof-packaging-only',
			evidenceNeeded: {
				honestBoundary: expect.stringContaining('No edit-producing rename'),
			},
		})
		expect(index.claimPortfolio[8]).toMatchObject({
			name: 'columnar-scan-sidecars',
			status: 'speculative-do-not-promote',
			handoffDecision: 'do-not-promote-yet',
			likelyHandoffOwner: ['performance'],
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
			acceptanceEvidence: expect.stringContaining(
				'representative streaming proofs covering passthrough/regenerate/add/drop',
			),
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
		expect(
			index.readiness.claimBlockerBoard.map((row) => `${row.artifact}/${row.ownerLoop}`),
		).toEqual([
			'safe-open-proof/performance',
			'safe-open-proof/product',
			'safe-open-proof/release',
			'package-action-proof/correctness',
			'package-action-proof/performance',
			'package-action-proof/product',
			'package-action-proof/release',
		])
		expect(index.readiness.claimBlockerBoard[0]).toMatchObject({
			artifact: 'safe-open-proof',
			claim: 'safe unknown workbook opening',
			ownerLoop: 'performance',
			blockerCount: 1,
			requirementIds: ['release-latency-run'],
			nextStepKinds: ['validation-run'],
			boundary: expect.stringContaining('routing evidence only'),
		})
		expect(
			index.readiness.claimBlockerBoard.find(
				(row) => row.artifact === 'package-action-proof' && row.ownerLoop === 'release',
			),
		).toMatchObject({
			blockerCount: 2,
			requirementIds: ['provenance-boundary', 'compact-report-publication-policy'],
			nextStepKinds: ['publication-policy'],
			forbiddenShortcuts: expect.arrayContaining([
				expect.stringContaining('signed provenance'),
				expect.stringContaining('compact report digests'),
			]),
		})
		expect(index.readiness.implementationHandoffs).toHaveLength(2)
		const safeOpenHandoff = index.readiness.implementationHandoffs[0]
		expect(safeOpenHandoff.rank).toBe(1)
		expect(safeOpenHandoff.artifact).toBe('safe-open-proof')
		expect(safeOpenHandoff.claim).toBe('safe unknown workbook opening')
		expect(safeOpenHandoff.ownerLoops).toEqual(['performance', 'product', 'release'])
		expect(safeOpenHandoff.implementationSurfacePromotionAllowed).toBe(false)
		expect(safeOpenHandoff.proofRequired.fixture).toContain('generated signed and malformed cases')
		expect(safeOpenHandoff.proofRequired.benchmark).toContain(
			'Release-environment open-plan latency',
		)
		expect(safeOpenHandoff.proofRequired.surface).toContain('no new opener surface')
		expect(safeOpenHandoff.proofRequired.validationGate).toContain('safe-open proof harness')
		expect(safeOpenHandoff.proofRequired.competitorContrast).toContain('Microsoft Protected View')
		expect(safeOpenHandoff.proofRequired.honestBoundary).toContain('Not malware scanning')
		expect(safeOpenHandoff.proofRequired.killCriterion).toContain(
			'Do not publish headline wording if generated signed or malformed packages are hidden',
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
			'generated signed/malformed',
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
			'generated signature topology',
		)
		expect(packageActionHandoff.blockingActions[1].forbiddenShortcut).toContain('chart XML')
		expect(packageActionHandoff.blockingActions[2].forbiddenShortcut).toContain(
			'full streaming parity',
		)
		expect(packageActionHandoff.blockingActions[3].acceptanceEvidence).toContain('Sigstore')
		expect(packageActionHandoff.blockingActions[4].forbiddenShortcut).toContain(
			'compact report digests',
		)
		expect(index.qssLeapfrogReleaseMatrix).toMatchObject({
			status: 'top-two-only',
			competitor: 'Quadratic/QSS',
			boundary: expect.stringContaining('top two claims only'),
		})
		expect(index.qssLeapfrogReleaseMatrix.rows).toHaveLength(2)
		expect(index.qssLeapfrogReleaseMatrix.rows.map((row) => row.artifact)).toEqual([
			'safe-open-proof',
			'package-action-proof',
		])
		expect(index.qssLeapfrogReleaseMatrix.rows[0]).toMatchObject({
			claim: 'safe unknown workbook opening',
		})
		expect(index.qssLeapfrogReleaseMatrix.rows[0].qssLikelyDoesWell).toEqual(
			expect.arrayContaining([expect.stringContaining('Python, SQL, JavaScript')]),
		)
		expect(index.qssLeapfrogReleaseMatrix.rows[0].ascendBetterWhereProven).toEqual(
			expect.arrayContaining([expect.stringContaining('Pre-hydration package-feature routing')]),
		)
		expect(index.qssLeapfrogReleaseMatrix.rows[0].claimsWeMustNotMake).toEqual(
			expect.arrayContaining([expect.stringContaining('Malware scanning')]),
		)
		expect(index.qssLeapfrogReleaseMatrix.rows[0].weakClaimDisposition).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ disposition: 'downgrade' }),
				expect.objectContaining({ disposition: 'blocker', ownerLoop: 'performance' }),
				expect.objectContaining({ disposition: 'kill', ownerLoop: 'release' }),
			]),
		)
		expect(index.qssLeapfrogReleaseMatrix.rows[0].acceptedEvidence).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					evidenceId: 'safe-open-proof-harness',
					command: 'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --json',
					path: 'fixtures/benchmarks/safe-open-proof.ts',
				}),
				expect.objectContaining({
					evidenceId: 'release-rc-gate',
					command: 'bun run release:rc:gate',
					path: 'scripts/release-rc-gate.ts',
				}),
			]),
		)
		expect(index.qssLeapfrogReleaseMatrix.rows[1]).toMatchObject({
			claim: 'auditable package-part mutation',
		})
		expect(index.qssLeapfrogReleaseMatrix.rows[1].ascendBetterWhereProven).toEqual(
			expect.arrayContaining([expect.stringContaining('Per-part write accounting')]),
		)
		expect(index.qssLeapfrogReleaseMatrix.rows[1].claimsWeMustNotMake).toEqual(
			expect.arrayContaining([expect.stringContaining('Signed provenance')]),
		)
		expect(index.qssLeapfrogReleaseMatrix.rows[1].acceptedEvidence).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					evidenceId: 'package-action-proof-harness',
					command: 'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
					path: 'fixtures/benchmarks/package-action-proof.ts',
				}),
				expect.objectContaining({
					evidenceId: 'copy-sheet-table-package-proof',
					kind: 'test',
					path: 'packages/sdk/src/agent-workflow.test.ts',
				}),
			]),
		)
		expect(index.qssLeapfrogReleaseMatrix.activeReleaseBlockers).toHaveLength(7)
		expect(index.qssLeapfrogReleaseMatrix.archivedResearchNotes.map((note) => note.name)).toEqual([
			'formula-language-service-primitives',
			'token-bounded-agent-view',
			'retained-viewport-patch-history',
			'release-proof-bundle',
			'formula-oracle-routing',
			'property-journal-laws',
			'columnar-scan-sidecars',
			'agent-workflow-observability',
			'research-surface-hygiene',
			'practical-latency-contracts',
		])
		expect(index.releaseDecisionBoard).toMatchObject({
			status: 'top-two-only',
			releaseGate: 'blocked-by-publication-policy',
			headlineClaimsAllowed: false,
			implementationSurfacePromotionAllowed: false,
			missingRequirementCount: 9,
			boundary: expect.stringContaining('Top-two release-decision artifact'),
		})
		expect(index.releaseDecisionBoard.rows.map((row) => row.artifact)).toEqual([
			'safe-open-proof',
			'package-action-proof',
		])
		for (const row of index.releaseDecisionBoard.rows) {
			expect(row.evidenceWeHave.length).toBeGreaterThan(0)
			expect(row.evidenceMissing.length).toBeGreaterThan(0)
			expect(row.qssContrast.length).toBeGreaterThan(0)
			expect(row.allowedWording.length).toBeGreaterThan(0)
			expect(row.forbiddenWording.length).toBeGreaterThan(0)
			expect(row.nextOwnerActions.length).toBeGreaterThan(0)
			expect(row.claimsWeMustNotMake.length).toBeGreaterThan(0)
			for (const ownerArtifact of row.ownerDecisionArtifacts) {
				expect(ownerArtifact.path).toMatch(/^(docs|fixtures|packages|scripts)\//)
				expect(ownerArtifact.validationCommand.length).toBeGreaterThan(0)
				expect(ownerArtifact.decision.length).toBeGreaterThan(0)
				expect(ownerArtifact.nextAction.length).toBeGreaterThan(0)
				expect(ownerArtifact.forbiddenShortcut.length).toBeGreaterThan(0)
			}
		}
		expect(
			index.releaseDecisionBoard.topClaimOwnerActionQueue.map(
				(row) =>
					`${row.ownerLoop}:${row.artifact}:${row.requirementId}:${row.nextStepKind}:${row.workBlockDisposition}`,
			),
		).toEqual([
			'product:safe-open-proof:public-edge-fixtures:owner-decision-or-fixture-replacement:benchmark-corpus-blocker',
			'performance:safe-open-proof:release-latency-run:validation-run:benchmark-corpus-blocker',
			'release:safe-open-proof:publication-boundary:publication-policy:implementation-ready-blocker',
			'release:safe-open-proof:compact-report-publication-policy:publication-policy:implementation-ready-blocker',
			'product:package-action-proof:edge-fixture-policy:owner-decision-or-fixture-replacement:benchmark-corpus-blocker',
			'correctness:package-action-proof:unsupported-feature-boundary:owner-boundary-approval:implementation-ready-blocker',
			'performance:package-action-proof:streaming-matrix-boundary:owner-decision-or-harness-expansion:benchmark-corpus-blocker',
			'release:package-action-proof:provenance-boundary:publication-policy:implementation-ready-blocker',
			'release:package-action-proof:compact-report-publication-policy:publication-policy:implementation-ready-blocker',
		])
		for (const row of index.releaseDecisionBoard.topClaimOwnerActionQueue) {
			expect(['implementation-ready-blocker', 'benchmark-corpus-blocker']).toContain(
				row.workBlockDisposition,
			)
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(Object)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.validationCommand.length).toBeGreaterThan(0)
			expect(row.acceptanceEvidence.length).toBeGreaterThan(0)
			expect(row.forbiddenShortcut.length).toBeGreaterThan(0)
			expect(row.allowedWording.length).toBeGreaterThan(0)
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction).toContain(row.validationCommand)
			expect(row.nextOwnerAction).toContain(row.acceptanceEvidence)
			expect(row.boundary).toContain('Owner-action queue row for top claims')
		}
		expect(index.releaseDecisionBoard.topClaimOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					ownerLoop: 'correctness',
					artifact: 'package-action-proof',
					requirementId: 'unsupported-feature-boundary',
					validationCommand:
						'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
					forbiddenShortcut: expect.stringContaining('chart XML'),
				}),
				expect.objectContaining({
					ownerLoop: 'performance',
					artifact: 'safe-open-proof',
					requirementId: 'release-latency-run',
					validationCommand:
						'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
				}),
			]),
		)
		expect(index.releaseDecisionBoard.doNotPromoteYet.map((item) => item.name)).toEqual([
			'formula-language-service-primitives',
			'token-bounded-agent-view',
			'retained-viewport-patch-history',
			'release-proof-bundle',
			'formula-oracle-routing',
			'property-journal-laws',
			'columnar-scan-sidecars',
			'agent-workflow-observability',
			'research-surface-hygiene',
			'practical-latency-contracts',
		])
		expect(
			index.releaseDecisionBoard.doNotPromoteYet.map(
				(item) => `${item.name}:${item.workBlockDisposition}`,
			),
		).toEqual([
			'formula-language-service-primitives:implementation-ready-blocker',
			'token-bounded-agent-view:implementation-ready-blocker',
			'retained-viewport-patch-history:implementation-ready-blocker',
			'release-proof-bundle:implementation-ready-blocker',
			'formula-oracle-routing:benchmark-corpus-blocker',
			'property-journal-laws:implementation-ready-blocker',
			'columnar-scan-sidecars:benchmark-corpus-blocker',
			'agent-workflow-observability:implementation-ready-blocker',
			'research-surface-hygiene:claim-downgrade-do-not-promote',
			'practical-latency-contracts:benchmark-corpus-blocker',
		])
		expect(index.releaseDecisionBoard.doNotPromoteDispositionSummary).toMatchObject({
			implementationReadyBlockerNames: [
				'formula-language-service-primitives',
				'token-bounded-agent-view',
				'retained-viewport-patch-history',
				'release-proof-bundle',
				'property-journal-laws',
				'agent-workflow-observability',
			],
			benchmarkCorpusBlockerNames: [
				'formula-oracle-routing',
				'columnar-scan-sidecars',
				'practical-latency-contracts',
			],
			claimDowngradeDoNotPromoteNames: ['research-surface-hygiene'],
			boundary: expect.stringContaining('Routing summary for blocked claims only'),
		})
		expect(index.releaseDecisionBoard.releaseWordingDecisionSummary).toMatchObject({
			status: 'headline-claims-blocked-local-wording-only',
			headlineClaimsAllowed: false,
			localAllowedClaimNames: ['safe-open-proof', 'package-action-proof'],
			doNotPromoteClaimNames: index.releaseDecisionBoard.doNotPromoteYet.map((item) => item.name),
			boundary: expect.stringContaining('Release wording decision summary only'),
		})
		expect(
			index.releaseDecisionBoard.releaseWordingDecisionSummary.localAllowedWordingByClaim[
				'safe-open-proof'
			],
		).toBe(index.releaseDecisionBoard.rows[0].allowedWording)
		expect(
			index.releaseDecisionBoard.releaseWordingDecisionSummary.doNotPromoteAllowedWordingByClaim[
				'research-surface-hygiene'
			],
		).toBe(
			index.releaseDecisionBoard.doNotPromoteYet.find(
				(item) => item.name === 'research-surface-hygiene',
			)?.allowedWording,
		)
		expect(
			index.releaseDecisionBoard.releaseWordingDecisionSummary.forbiddenWordingByClaim[
				'package-action-proof'
			],
		).toEqual(index.releaseDecisionBoard.rows[1].forbiddenWording)
		expect(index.releaseDecisionBoard.claimDecisionContractCoverage).toMatchObject({
			status: 'all-release-claim-decisions-self-contained',
			decisionCount: 12,
			topClaimDecisionCount: 2,
			doNotPromoteDecisionCount: 10,
			missingEvidenceWeHaveKeys: [],
			missingEvidenceMissingKeys: [],
			missingQssContrastKeys: [],
			missingAllowedWordingKeys: [],
			missingForbiddenWordingKeys: [],
			missingNextOwnerActionKeys: [],
			boundary: expect.stringContaining('Claim decision contract coverage only'),
		})
		expect(
			index.releaseDecisionBoard.blockedOwnerActionQueue.map(
				(row) => `${row.ownerLoop}:${row.name}:${row.workBlockDisposition}`,
			),
		).toEqual([
			'product:formula-language-service-primitives:implementation-ready-blocker',
			'correctness:formula-language-service-primitives:implementation-ready-blocker',
			'product:token-bounded-agent-view:implementation-ready-blocker',
			'product:retained-viewport-patch-history:implementation-ready-blocker',
			'performance:retained-viewport-patch-history:implementation-ready-blocker',
			'product:release-proof-bundle:implementation-ready-blocker',
			'release:release-proof-bundle:implementation-ready-blocker',
			'correctness:formula-oracle-routing:benchmark-corpus-blocker',
			'correctness:property-journal-laws:implementation-ready-blocker',
			'performance:columnar-scan-sidecars:benchmark-corpus-blocker',
			'product:agent-workflow-observability:implementation-ready-blocker',
			'product:research-surface-hygiene:claim-downgrade-do-not-promote',
			'release:research-surface-hygiene:claim-downgrade-do-not-promote',
			'performance:practical-latency-contracts:benchmark-corpus-blocker',
		])
		for (const row of index.releaseDecisionBoard.blockedOwnerActionQueue) {
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.allowedWording).toContain('Do not promote')
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction.length).toBeGreaterThan(0)
			expect(row.validationCommands).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.validationCommands.every((command) => command.length > 0)).toBe(true)
			expect(row.boundary).toContain('Owner-action queue row')
		}
		expect(
			index.releaseDecisionBoard.blockedOwnerActionQueue
				.filter((row) => row.workBlockDisposition === 'benchmark-corpus-blocker')
				.map((row) => row.name),
		).toEqual(['formula-oracle-routing', 'columnar-scan-sidecars', 'practical-latency-contracts'])
		expect(
			index.releaseDecisionBoard.benchmarkCorpusOwnerActionQueue.map(
				(row) => `${row.sourceQueue}:${row.ownerLoop}:${row.name}:${row.requirementId ?? 'none'}`,
			),
		).toEqual([
			'top-claim-owner-action:product:safe-open-proof:public-edge-fixtures',
			'top-claim-owner-action:performance:safe-open-proof:release-latency-run',
			'top-claim-owner-action:product:package-action-proof:edge-fixture-policy',
			'top-claim-owner-action:performance:package-action-proof:streaming-matrix-boundary',
			'blocked-owner-action:correctness:formula-oracle-routing:none',
			'blocked-owner-action:performance:columnar-scan-sidecars:none',
			'blocked-owner-action:performance:practical-latency-contracts:none',
		])
		for (const row of index.releaseDecisionBoard.benchmarkCorpusOwnerActionQueue) {
			expect(row.workBlockDisposition).toBe('benchmark-corpus-blocker')
			expect(row.validationCommands).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.allowedWording.length).toBeGreaterThan(0)
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction.length).toBeGreaterThan(0)
			expect(row.boundary).toContain('Benchmark/corpus owner-action queue row')
		}
		expect(index.releaseDecisionBoard.benchmarkCorpusOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					sourceQueue: 'top-claim-owner-action',
					name: 'safe-open-proof',
					requirementId: 'release-latency-run',
					validationCommands: [
						'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
					],
				}),
				expect.objectContaining({
					sourceQueue: 'blocked-owner-action',
					name: 'practical-latency-contracts',
					validationCommands: expect.arrayContaining([
						expect.stringContaining(
							'--input-preset public-tracked --contract all --repeat 3 --warmup 1 --json',
						),
					]),
				}),
			]),
		)
		expect(
			index.releaseDecisionBoard.implementationReadyOwnerActionQueue.map(
				(row) => `${row.sourceQueue}:${row.ownerLoop}:${row.name}:${row.requirementId ?? 'none'}`,
			),
		).toEqual([
			'top-claim-owner-action:release:safe-open-proof:publication-boundary',
			'top-claim-owner-action:release:safe-open-proof:compact-report-publication-policy',
			'top-claim-owner-action:correctness:package-action-proof:unsupported-feature-boundary',
			'top-claim-owner-action:release:package-action-proof:provenance-boundary',
			'top-claim-owner-action:release:package-action-proof:compact-report-publication-policy',
			'blocked-owner-action:product:formula-language-service-primitives:none',
			'blocked-owner-action:correctness:formula-language-service-primitives:none',
			'blocked-owner-action:product:token-bounded-agent-view:none',
			'blocked-owner-action:product:retained-viewport-patch-history:none',
			'blocked-owner-action:performance:retained-viewport-patch-history:none',
			'blocked-owner-action:product:release-proof-bundle:none',
			'blocked-owner-action:release:release-proof-bundle:none',
			'blocked-owner-action:correctness:property-journal-laws:none',
			'blocked-owner-action:product:agent-workflow-observability:none',
		])
		for (const row of index.releaseDecisionBoard.implementationReadyOwnerActionQueue) {
			expect(row.workBlockDisposition).toBe('implementation-ready-blocker')
			expect(row.validationCommands).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.allowedWording.length).toBeGreaterThan(0)
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction.length).toBeGreaterThan(0)
			expect(row.boundary).toContain('Implementation-ready owner-action queue row')
		}
		expect(index.releaseDecisionBoard.implementationReadyOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					sourceQueue: 'top-claim-owner-action',
					name: 'package-action-proof',
					requirementId: 'unsupported-feature-boundary',
					validationCommands: [
						'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
					],
				}),
				expect.objectContaining({
					sourceQueue: 'blocked-owner-action',
					name: 'formula-language-service-primitives',
					validationCommands: expect.arrayContaining([
						expect.stringContaining('formula-assist-proof.ts'),
					]),
				}),
			]),
		)
		expect(
			index.releaseDecisionBoard.claimDowngradeOwnerActionQueue.map(
				(row) => `${row.sourceQueue}:${row.ownerLoop}:${row.name}`,
			),
		).toEqual([
			'blocked-owner-action:product:research-surface-hygiene',
			'blocked-owner-action:release:research-surface-hygiene',
		])
		for (const row of index.releaseDecisionBoard.claimDowngradeOwnerActionQueue) {
			expect(row.workBlockDisposition).toBe('claim-downgrade-do-not-promote')
			expect(row.ownerFiles).toEqual(
				expect.arrayContaining([
					'research/',
					'scripts/ascend-loop-manager.ts',
					'tmp/ascend-loop-manager/',
				]),
			)
			expect(row.validationCommands).toEqual([
				'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
				'bun test fixtures/benchmarks/release-proof-index.test.ts',
			])
			expect(row.commandsToRun).toEqual(row.validationCommands)
			expect(row.failureEvidence).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.acceptanceCriteria).toContain('classifies each untriaged path')
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.allowedWording).toContain('Do not promote research-surface-hygiene')
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction).toContain(
				'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
			)
			expect(row.boundary).toContain('Claim downgrade owner-action queue row')
		}
		expect(index.releaseDecisionBoard.ownerActionQueueCoverage).toMatchObject({
			status: 'all-owner-actions-covered-by-disposition-queues',
			sourceTopClaimActionCount: 9,
			sourceBlockedActionCount: 14,
			benchmarkCorpusActionCount: 7,
			implementationReadyActionCount: 14,
			claimDowngradeActionCount: 2,
			coveredActionCount: 23,
			uncoveredTopClaimActionKeys: [],
			uncoveredBlockedActionKeys: [],
			boundary: expect.stringContaining('Owner action queue coverage only'),
		})
		expect(index.releaseDecisionBoard.ownerActionExecutionContractCoverage).toMatchObject({
			status: 'all-disposition-owner-actions-have-execution-contract',
			actionCount: 23,
			benchmarkCorpusActionCount: 7,
			implementationReadyActionCount: 14,
			claimDowngradeActionCount: 2,
			missingOwnerFileActionKeys: [],
			missingCommandActionKeys: [],
			missingFailureEvidenceActionKeys: [],
			missingAcceptanceCriteriaActionKeys: [],
			boundary: expect.stringContaining('Execution contract coverage only'),
		})
		expect(
			index.releaseDecisionBoard.doNotPromoteYet.every(
				(item) => item.status === 'do-not-promote-yet',
			),
		).toBe(true)
		expect(index.releaseDecisionBoard.doNotPromoteYet[0].boundary).toContain(
			'Do not turn this into release wording',
		)
		const formulaLanguageDecision = index.releaseDecisionBoard.doNotPromoteYet.find(
			(item) => item.name === 'formula-language-service-primitives',
		)
		const formulaLanguageEvidence = formulaLanguageDecision?.evidenceWeHave.join('\n') ?? ''
		const formulaLanguageMissing = formulaLanguageDecision?.evidenceMissing.join('\n') ?? ''
		const formulaLanguageForbidden = formulaLanguageDecision?.forbiddenWording.join('\n') ?? ''
		const formulaLanguageNextOwnerAction = formulaLanguageDecision?.nextOwnerAction ?? ''
		expect(formulaLanguageEvidence).toContain('formula-assist-proof.ts')
		expect(formulaLanguageEvidence).toContain('packages/sdk/src/formula-edit.test.ts')
		expect(formulaLanguageEvidence).toContain('apps/cli/src/cli.test.ts')
		expect(formulaLanguageEvidence).toContain('apps/api/src/server.test.ts')
		expect(formulaLanguageEvidence).toContain('apps/mcp/src/index.test.ts')
		expect(formulaLanguageMissing).toContain('public formula corpus')
		expect(formulaLanguageMissing).toContain('operation-owned edit plans')
		expect(formulaLanguageForbidden).toContain('edit-producing rename')
		expect(formulaLanguageForbidden).toContain('external-ref rename')
		expect(formulaLanguageNextOwnerAction).toContain('SDK/CLI/API/MCP')
		expect(formulaLanguageNextOwnerAction).toContain('formula-assist-proof.ts')
		expect(formulaLanguageDecision?.validationCommands).toEqual(
			expect.arrayContaining([
				'bun run fixtures/benchmarks/formula-assist-proof.ts --sample 250 --no-timings --json',
				expect.stringContaining('packages/sdk/src/formula-edit.test.ts'),
			]),
		)
		const researchSurfaceDecision = index.releaseDecisionBoard.doNotPromoteYet.find(
			(item) => item.name === 'research-surface-hygiene',
		)
		expect(researchSurfaceDecision?.ownerLoops).toEqual(['product', 'release'])
		expect(researchSurfaceDecision?.evidenceMissing.join('\n')).toContain(
			'Classify current research files',
		)
		expect(researchSurfaceDecision?.evidenceMissing.join('\n')).toContain(
			'Inventory of current research files',
		)
		expect(researchSurfaceDecision?.evidenceWeHave.join('\n')).toContain(
			'scripts/ascend-loop-manager.ts',
		)
		expect(researchSurfaceDecision?.evidenceMissing.join('\n')).toContain(
			'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
		)
		expect(researchSurfaceDecision?.allowedWording).toContain(
			'Do not promote research-surface-hygiene',
		)
		expect(researchSurfaceDecision?.forbiddenWording.join('\n')).toContain(
			'Do not promote any research-derived claim',
		)
		expect(researchSurfaceDecision?.forbiddenWording.join('\n')).toContain(
			'Untriaged research files are not release evidence',
		)
		expect(researchSurfaceDecision?.forbiddenWording.join('\n')).toContain(
			'Do not cite `research/` or `tmp/` files',
		)
		expect(researchSurfaceDecision?.nextOwnerAction).toContain(
			'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
		)
		expect(researchSurfaceDecision?.nextOwnerAction).toContain(
			'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
		)
		expect(researchSurfaceDecision?.nextOwnerAction).toContain(
			'bun test fixtures/benchmarks/release-proof-index.test.ts',
		)
		expect(researchSurfaceDecision?.validationCommands).toEqual([
			'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
			'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
			'bun test fixtures/benchmarks/release-proof-index.test.ts',
		])
		const tokenBoundedDecision = index.releaseDecisionBoard.doNotPromoteYet.find(
			(item) => item.name === 'token-bounded-agent-view',
		)
		const tokenBoundedEvidence = tokenBoundedDecision?.evidenceWeHave.join('\n') ?? ''
		const tokenBoundedMissing = tokenBoundedDecision?.evidenceMissing.join('\n') ?? ''
		const tokenBoundedForbidden = tokenBoundedDecision?.forbiddenWording.join('\n') ?? ''
		const tokenBoundedNextOwnerAction = tokenBoundedDecision?.nextOwnerAction ?? ''
		expect(tokenBoundedEvidence).toContain('agent-view-budget-proof.test.ts')
		expect(tokenBoundedEvidence).toContain('agent-view-recovery-proof.test.ts')
		expect(tokenBoundedEvidence).toContain('apps/mcp/src/index.test.ts')
		expect(tokenBoundedMissing).toContain('One public product example')
		expect(tokenBoundedMissing).toContain('approximate token estimates')
		expect(tokenBoundedMissing).toContain('SDK/CLI/API/MCP agent-view budget tests')
		expect(tokenBoundedForbidden).toContain('exact model-token counts')
		expect(tokenBoundedNextOwnerAction).toContain('maxApproxTokens')
		expect(tokenBoundedNextOwnerAction).toEqual(
			expect.stringContaining(
				'bun test fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts',
			),
		)
		expect(tokenBoundedDecision?.validationCommands).toEqual(
			expect.arrayContaining([
				'bun test fixtures/benchmarks/agent-view-budget-proof.test.ts fixtures/benchmarks/agent-view-recovery-proof.test.ts',
				expect.stringContaining('packages/sdk/src/sdk.test.ts'),
			]),
		)
		const viewportPatchDecision = index.releaseDecisionBoard.doNotPromoteYet.find(
			(item) => item.name === 'retained-viewport-patch-history',
		)
		const viewportPatchEvidence = viewportPatchDecision?.evidenceWeHave.join('\n') ?? ''
		const viewportPatchMissing = viewportPatchDecision?.evidenceMissing.join('\n') ?? ''
		const viewportPatchForbidden = viewportPatchDecision?.forbiddenWording.join('\n') ?? ''
		const viewportPatchNextOwnerAction = viewportPatchDecision?.nextOwnerAction ?? ''
		expect(viewportPatchEvidence).toContain('viewport-patch-proof.test.ts')
		expect(viewportPatchEvidence).toContain('packages/sdk/src/interactive-contract.test.ts')
		expect(viewportPatchEvidence).toContain('apps/api/src/server.test.ts')
		expect(viewportPatchEvidence).toContain('apps/mcp/src/index.test.ts')
		expect(viewportPatchMissing).toContain('patch bytes')
		expect(viewportPatchMissing).toContain('retention caps')
		expect(viewportPatchForbidden).toContain('CRDT')
		expect(viewportPatchForbidden).toContain('unlimited history')
		expect(viewportPatchNextOwnerAction).toContain('SDK/API/MCP')
		expect(viewportPatchNextOwnerAction).toContain(
			'bun test fixtures/benchmarks/viewport-patch-proof.test.ts',
		)
		expect(viewportPatchDecision?.validationCommands).toEqual(
			expect.arrayContaining([
				'bun test fixtures/benchmarks/viewport-patch-proof.test.ts',
				expect.stringContaining('packages/sdk/src/interactive-contract.test.ts'),
			]),
		)
		for (const item of index.releaseDecisionBoard.doNotPromoteYet) {
			const qssContrast = item.qssContrast.join('\n')
			expect([
				'implementation-ready-blocker',
				'benchmark-corpus-blocker',
				'claim-downgrade-do-not-promote',
			]).toContain(item.workBlockDisposition)
			expect(item.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(qssContrast).not.toContain(
				'QSS contrast is blocked until this diagnostic evidence changes a top-two release claim.',
			)
			expect(item.allowedWording).toContain('Do not promote')
			expect(item.allowedWording).not.toContain('owner planning or research evidence')
			expect(item.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.nextOwnerAction.length).toBeGreaterThan(0)
			expect(item.validationCommands).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.validationCommands.every((command) => command.length > 0)).toBe(true)
			expect(item.nextOwnerAction).not.toContain(
				'No owner action is release-blocking until this claim changes the top-two release gate.',
			)
		}
		expect(index.releaseDecisionBoard.doNotPromoteYet[0]).toMatchObject({
			evidenceWeHave: expect.arrayContaining([
				expect.stringContaining('rejection-first prepareRename'),
				expect.stringContaining('formula-assist-proof'),
			]),
			evidenceMissing: expect.arrayContaining([
				expect.stringContaining('Workbook-context ownership'),
				expect.stringContaining('Public formula corpus'),
			]),
			qssContrast: expect.arrayContaining([expect.stringContaining('HyperFormula')]),
			forbiddenWording: expect.arrayContaining([
				expect.stringContaining('Do not promote rename'),
				expect.stringContaining('No edit-producing rename'),
			]),
			nextOwnerAction: expect.stringContaining('Workbook-context ownership'),
		})
		expect(
			index.releaseDecisionBoard.doNotPromoteYet.find(
				(item) => item.name === 'release-proof-bundle',
			),
		).toMatchObject({
			evidenceWeHave: expect.arrayContaining([
				expect.stringContaining('bun run release:rc:gate'),
				expect.stringContaining('safe-open-proof.ts --no-timings --compact-json'),
				expect.stringContaining('package-action-proof.ts --no-timings --compact-json'),
			]),
			evidenceMissing: expect.arrayContaining([
				expect.stringContaining('One real public workbook workflow per top claim'),
				expect.stringContaining('artifact storage path'),
				expect.stringContaining('Passing validation sequence'),
				expect.stringContaining('Golden proof fixtures'),
			]),
			qssContrast: expect.arrayContaining([
				expect.stringContaining('Generic spreadsheet libraries'),
			]),
			forbiddenWording: expect.arrayContaining([
				expect.stringContaining('third-party attestation'),
				expect.stringContaining('Not signed'),
			]),
			validationCommands: expect.arrayContaining([
				'bun run release:rc:gate',
				'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json',
				'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --compact-json',
			]),
			nextOwnerAction: expect.stringContaining('bun run release:rc:gate'),
		})
		const formulaOracleDecision = index.releaseDecisionBoard.doNotPromoteYet.find(
			(item) => item.name === 'formula-oracle-routing',
		)
		const formulaOracleEvidence = formulaOracleDecision?.evidenceWeHave.join('\n') ?? ''
		const formulaOracleMissing = formulaOracleDecision?.evidenceMissing.join('\n') ?? ''
		const formulaOracleForbidden = formulaOracleDecision?.forbiddenWording.join('\n') ?? ''
		const formulaOracleNextOwnerAction = formulaOracleDecision?.nextOwnerAction ?? ''
		expect(formulaOracleEvidence).toContain('formula-corpus-correctness.test.ts')
		expect(formulaOracleEvidence).toContain('fixtures/xlsx/libreoffice/manifest.ts')
		expect(formulaOracleEvidence).toContain('formula-sota.test.ts')
		expect(formulaOracleMissing).toContain('unsupported function')
		expect(formulaOracleMissing).toContain('HyperFormula, LibreOffice, Excel')
		expect(formulaOracleMissing).toContain('artifact verifier')
		expect(formulaOracleForbidden).toContain('Excel-compatible formulas')
		expect(formulaOracleForbidden).toContain('fresh cached values')
		expect(formulaOracleForbidden).toContain('QSS/SOTA formula superiority')
		expect(formulaOracleNextOwnerAction).toContain('formula-corpus-correctness.test.ts')
		expect(formulaOracleNextOwnerAction).toContain('--tag formula-fidelity')
		expect(formulaOracleDecision?.validationCommands).toEqual(
			expect.arrayContaining([
				'bun test fixtures/benchmarks/formula-corpus-correctness.test.ts --timeout 30000',
				expect.stringContaining('--tag formula-fidelity'),
			]),
		)
		const propertyJournalDecision = index.releaseDecisionBoard.doNotPromoteYet.find(
			(item) => item.name === 'property-journal-laws',
		)
		const propertyJournalEvidence = propertyJournalDecision?.evidenceWeHave.join('\n') ?? ''
		const propertyJournalMissing = propertyJournalDecision?.evidenceMissing.join('\n') ?? ''
		const propertyJournalForbidden = propertyJournalDecision?.forbiddenWording.join('\n') ?? ''
		const propertyJournalNextOwnerAction = propertyJournalDecision?.nextOwnerAction ?? ''
		expect(propertyJournalDecision?.allowedWording).toContain(
			'Do not promote property-journal-laws',
		)
		expect(propertyJournalEvidence).toContain('journal-law-proof.test.ts')
		expect(propertyJournalEvidence).toContain('packages/sdk/src/journal-exactness.test.ts')
		expect(propertyJournalEvidence).toContain('deterministic local journal evidence')
		expect(propertyJournalMissing).toContain('Generated operation sequences')
		expect(propertyJournalMissing).toContain('Shrinkable and replayable property generation')
		expect(propertyJournalMissing).toContain('style exactness')
		expect(propertyJournalForbidden).toContain('Do not promote broad inverse-law claims')
		expect(propertyJournalForbidden).toContain('full undo coverage')
		expect(propertyJournalForbidden).toContain('signed audit')
		expect(propertyJournalNextOwnerAction).toContain('fast-check')
		expect(propertyJournalNextOwnerAction).toContain('journal-law-proof.test.ts')
		expect(propertyJournalDecision?.validationCommands).toEqual([
			'bun test fixtures/benchmarks/journal-law-proof.test.ts --timeout 30000',
			'bun test packages/sdk/src/journal-exactness.test.ts --timeout 30000',
		])
		const columnarSidecarDecision = index.releaseDecisionBoard.doNotPromoteYet.find(
			(item) => item.name === 'columnar-scan-sidecars',
		)
		const columnarSidecarEvidence = columnarSidecarDecision?.evidenceWeHave.join('\n') ?? ''
		const columnarSidecarMissing = columnarSidecarDecision?.evidenceMissing.join('\n') ?? ''
		const columnarSidecarForbidden = columnarSidecarDecision?.forbiddenWording.join('\n') ?? ''
		const columnarSidecarNextOwnerAction = columnarSidecarDecision?.nextOwnerAction ?? ''
		expect(columnarSidecarEvidence).toContain('columnar-sidecar.test.ts')
		expect(columnarSidecarEvidence).toContain('sec-mmf-statistics-2022-02.xlsx')
		expect(columnarSidecarEvidence).toContain('workbook grid as source of truth')
		expect(columnarSidecarMissing).toContain('structurally diverse external public workbook')
		expect(columnarSidecarMissing).toContain('memory caps')
		expect(columnarSidecarMissing).toContain('generation-key invalidation')
		expect(columnarSidecarForbidden).toContain('Arrow ABI')
		expect(columnarSidecarForbidden).toContain('QSS/SOTA speed win')
		expect(columnarSidecarForbidden).toContain('SDK/API/MCP sidecar product surface')
		expect(columnarSidecarNextOwnerAction).toContain('columnar-sidecar.test.ts')
		expect(columnarSidecarNextOwnerAction).toContain('claim-report --json')
		expect(columnarSidecarDecision?.allowedWording).toContain('benchmark-only disposable sidecar')
		expect(columnarSidecarDecision?.validationCommands).toEqual([
			'bun test fixtures/benchmarks/columnar-sidecar.test.ts --timeout 30000',
			'bun run fixtures/benchmarks/columnar-sidecar.ts --fixture fixtures/xlsx/external/sec-mmf-statistics-2022-02.xlsx --sheet "Table 9" --repeats 8 --claim-report --json',
		])
		const agentWorkflowDecision = index.releaseDecisionBoard.doNotPromoteYet.find(
			(item) => item.name === 'agent-workflow-observability',
		)
		const agentWorkflowEvidence = agentWorkflowDecision?.evidenceWeHave.join('\n') ?? ''
		const agentWorkflowMissing = agentWorkflowDecision?.evidenceMissing.join('\n') ?? ''
		const agentWorkflowForbidden = agentWorkflowDecision?.forbiddenWording.join('\n') ?? ''
		const agentWorkflowNextOwnerAction = agentWorkflowDecision?.nextOwnerAction ?? ''
		expect(agentWorkflowEvidence).toContain('packages/sdk/src/agent-workflow.test.ts')
		expect(agentWorkflowEvidence).toContain('apps/cli/src/cli.test.ts')
		expect(agentWorkflowEvidence).toContain('apps/api/src/server.test.ts')
		expect(agentWorkflowEvidence).toContain('apps/mcp/src/index.test.ts')
		expect(agentWorkflowEvidence).toContain('docs/AGENT_WORKFLOW.md')
		expect(agentWorkflowMissing).toContain('inspect, plan, commit, reopen, diff, audit')
		expect(agentWorkflowMissing).toContain('Trace payload size')
		expect(agentWorkflowMissing).toContain('Golden trace fixtures')
		expect(agentWorkflowForbidden).toContain('autonomous correctness')
		expect(agentWorkflowForbidden).toContain('signed audit trail')
		expect(agentWorkflowForbidden).toContain('repair automation')
		expect(agentWorkflowDecision?.allowedWording).toContain('internal workflow evidence')
		expect(agentWorkflowNextOwnerAction).toContain('failure taxonomy')
		expect(agentWorkflowNextOwnerAction).toContain('trace payload size')
		expect(agentWorkflowNextOwnerAction).toContain('packages/sdk/src/agent-workflow.test.ts')
		expect(agentWorkflowNextOwnerAction).toContain('apps/cli/src/cli.test.ts')
		expect(agentWorkflowNextOwnerAction).toContain('apps/api/src/server.test.ts')
		expect(agentWorkflowNextOwnerAction).toContain('apps/mcp/src/index.test.ts')
		expect(agentWorkflowNextOwnerAction).toContain('release-proof-index.ts --no-timings')
		expect(agentWorkflowDecision?.validationCommands).toEqual(
			expect.arrayContaining([
				expect.stringContaining('packages/sdk/src/agent-workflow.test.ts'),
				expect.stringContaining('apps/cli/src/cli.test.ts'),
				expect.stringContaining('apps/api/src/server.test.ts'),
				expect.stringContaining('apps/mcp/src/index.test.ts'),
				'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
			]),
		)
		const practicalLatencyDecision = index.releaseDecisionBoard.doNotPromoteYet.at(-1)
		const practicalLatencyNextOwnerAction = String(practicalLatencyDecision?.nextOwnerAction ?? '')
		expect(practicalLatencyDecision).toMatchObject({
			name: 'practical-latency-contracts',
			allowedWording: expect.stringContaining('benchmark-owner diagnostics'),
			evidenceWeHave: expect.arrayContaining([
				expect.stringContaining('practical-latency-contracts.ts'),
				expect.stringContaining('practical-latency-contracts.test.ts'),
			]),
			evidenceMissing: expect.arrayContaining([
				expect.stringContaining('tracked-clean run over standardized public inputs'),
				expect.stringContaining('--input-preset public-tracked --contract all --repeat 3'),
				expect.stringContaining('summary/profile artifacts'),
			]),
			qssContrast: expect.arrayContaining([
				expect.stringContaining('visible spreadsheet responsiveness'),
			]),
			forbiddenWording: expect.arrayContaining([
				expect.stringContaining('No local timing report'),
				expect.stringContaining('fastest XLSX reader'),
				expect.stringContaining('hot-cache-only timing evidence'),
			]),
			nextOwnerAction: expect.stringContaining(
				'bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract all --repeat 3 --warmup 1 --json',
			),
			validationCommands: [
				'bun test fixtures/benchmarks/practical-latency-contracts.test.ts --timeout 30000',
				'bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract all --repeat 1 --warmup 0 --dry-run --json',
				'bun run fixtures/benchmarks/practical-latency-contracts.ts --input-preset public-tracked --contract all --repeat 3 --warmup 1 --json',
			],
		})
		expect(practicalLatencyNextOwnerAction).toContain(
			'bun test fixtures/benchmarks/practical-latency-contracts.test.ts --timeout 30000',
		)
		expect(practicalLatencyNextOwnerAction).toContain('--dry-run --json')
		const safeOpenDecision = index.releaseDecisionBoard.rows[0]
		expect(safeOpenDecision.claimWordingAllowedToday).toBe('safe unknown workbook opening')
		expect(safeOpenDecision.evidenceWeHave.map((item) => item.evidenceId)).toContain(
			'safe-open-proof-harness',
		)
		expect(safeOpenDecision.evidenceMissing.join('\n')).toContain('public-edge-fixtures')
		expect(safeOpenDecision.qssContrast.join('\n')).toContain('QSS likely does well')
		expect(safeOpenDecision.qssContrast.join('\n')).toContain('Ascend proven today')
		expect(safeOpenDecision.allowedWording).toContain('pre-hydration package-feature routing')
		expect(safeOpenDecision.forbiddenWording.join('\n')).toContain('Malware scanning')
		expect(safeOpenDecision.nextOwnerActions.map((action) => action.requirementId)).toContain(
			'release-latency-run',
		)
		expect(safeOpenDecision.ownerDecisionArtifacts).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					ownerLoop: 'correctness',
					artifactId: 'excel-behavior-compatibility-matrix',
					path: 'docs/EXCEL_BEHAVIOR_COMPATIBILITY_MATRIX.md',
					validationCommand:
						'bun test packages/sdk/src/excel-behavior-compatibility-matrix.test.ts',
					decision: expect.stringContaining('protection metadata'),
					nextAction: expect.stringContaining('public LibreOffice cached-value parity'),
					forbiddenShortcut: expect.stringContaining('full chart editing support'),
				}),
				expect.objectContaining({
					ownerLoop: 'performance',
					artifactId: 'performance-claim-baseline-matrix',
					path: 'docs/PERFORMANCE_CLAIM_BASELINE_MATRIX.md',
					validationCommand:
						'bun test fixtures/benchmarks/performance-claim-baseline-matrix.test.ts',
					nextAction: expect.stringContaining('focused ClosedXML head-to-head read run'),
					decision: expect.stringContaining('performance matrix as a defer decision'),
					forbiddenShortcut: expect.stringContaining('one-workload medians'),
				}),
			]),
		)
		expect(
			safeOpenDecision.nextOwnerActions.find(
				(action) => action.requirementId === 'release-latency-run',
			)?.validationCommand,
		).toBe('bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json')
		expect(safeOpenDecision.headlineClaimAllowed).toBe(false)
		expect(safeOpenDecision.implementationSurfacePromotionAllowed).toBe(false)
		expect(safeOpenDecision.proofRequired.fixture).toContain('Public clean')
		expect(safeOpenDecision.proofRequired.validationGate).toContain('safe-open proof harness')
		expect(safeOpenDecision.acceptedEvidence.map((item) => item.evidenceId)).toContain(
			'safe-open-proof-harness',
		)
		expect(safeOpenDecision.claimsWeMustNotMake.join('\n')).toContain('Malware scanning')
		expect(
			safeOpenDecision.aPlusBlockingOwnerActions.map(
				(action) => `${action.requirementId}/${action.ownerLoop}`,
			),
		).toContain('release-latency-run/performance')
		const packageActionDecision = index.releaseDecisionBoard.rows[1]
		expect(packageActionDecision.claimWordingAllowedToday).toBe('auditable package-part mutation')
		expect(packageActionDecision.evidenceMissing.join('\n')).toContain('edge-fixture-policy')
		expect(packageActionDecision.allowedWording).toContain('per-part package action accounting')
		expect(packageActionDecision.ownerDecisionArtifacts).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					ownerLoop: 'correctness',
					artifactId: 'excel-behavior-compatibility-matrix',
					path: 'docs/EXCEL_BEHAVIOR_COMPATIBILITY_MATRIX.md',
					decision: expect.stringContaining('protection metadata'),
					forbiddenShortcut: expect.stringContaining(
						'worksheet/workbook/protected-range protection as file security',
					),
				}),
				expect.objectContaining({
					ownerLoop: 'performance',
					artifactId: 'package-action-streaming-matrix-evidence',
					path: 'fixtures/benchmarks/package-action-proof.ts',
					validationCommand:
						'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
					nextAction: expect.stringContaining('representative passthrough/regenerate/add/drop'),
					forbiddenShortcut: expect.stringContaining('full streaming writer parity'),
				}),
			]),
		)
		expect(packageActionDecision.nextOwnerActions.map((action) => action.requirementId)).toContain(
			'streaming-matrix-boundary',
		)
		expect(
			packageActionDecision.nextOwnerActions.find(
				(action) => action.requirementId === 'streaming-matrix-boundary',
			)?.validationCommand,
		).toBe('bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json')
		expect(packageActionDecision.proofRequired.honestBoundary).toContain('Not signed provenance')
		expect(packageActionDecision.claimsWeMustNotMake.join('\n')).toContain('Signed provenance')
		expect(packageActionDecision.claimsWeMustNotMake.join('\n')).toContain('Full streaming parity')
		expect(
			packageActionDecision.aPlusBlockingOwnerActions.map(
				(action) => `${action.requirementId}/${action.ownerLoop}`,
			),
		).toContain('streaming-matrix-boundary/performance')
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
				publicFixtureCases: 10,
				generatedWorkbookCases: 0,
				generatedEdgePackageCases: 1,
				malformedCases: 1,
				generatedCaseNames: ['signed', 'malformed'],
				deterministicGeneratedCaseNames: ['signed', 'malformed'],
				generatedCaseSha256: {
					signed: expect.stringMatching(/^[a-f0-9]{64}$/),
					malformed: expect.stringMatching(/^[a-f0-9]{64}$/),
				},
			},
			summary: { cases: 12, rejected: 3, malformedRejected: true },
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
				publicFixtureCases: 5,
				generatedWorkbookCases: 2,
				generatedEdgePackageCases: 1,
				malformedCases: 0,
				generatedCaseNames: [
					'regenerate-existing-sheet',
					'add-sheet-part',
					'signature-invalidation-drop',
				],
				deterministicGeneratedCaseNames: ['signature-invalidation-drop'],
				generatedCaseSha256: expect.objectContaining({
					'signature-invalidation-drop': expect.stringMatching(/^[a-f0-9]{64}$/),
				}),
			},
			summary: {
				cases: 8,
				allActionsCovered: true,
				sourceGraphEverywhere: true,
				streamingProofCases: 5,
				streamingRegenerateParts: 4,
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
				'safe-open-proof': ['signed', 'malformed'],
				'package-action-proof': ['signature-invalidation-drop'],
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
		expect(handoff.fixturePolicyEvidence.safeOpen.signatureOrUnknownMatches).toBe(1)
		expect(handoff.fixturePolicyEvidence.safeOpen.rejectedFixtures).toEqual([
			'fixtures/xlsx/calamine/pass_protected.xlsx',
		])
		expect(handoff.fixturePolicyEvidence.packageAction.missingReplacementFeatures).toEqual([
			'signaturePackage',
		])
		expect(handoff.generatedFixtureDecisionEvidence).toMatchObject({
			status: 'generated-structural-cases-disclosed-owner-approval-required',
			allGeneratedStructuralCasesDisclosed: true,
			publicReplacementGapsRemain: true,
			ownerApprovalRequired: true,
		})
		expect(handoff.generatedFixtureDecisionEvidence.cases.map((entry) => entry.caseName)).toEqual([
			'signed',
			'malformed',
			'signature-invalidation-drop',
		])
		expect(handoff.performancePolicy.approvalChecklist.map((item) => item.gateId)).toEqual([
			'release-latency-run',
			'streaming-matrix-boundary',
		])
		expect(handoff.safeOpenLatencyValidationEvidence).toMatchObject({
			status: 'timed-evidence-absent-owner-run-required',
			timedCaseCount: 0,
			ownerApprovalRequired: true,
			releaseClaimAllowed: false,
			thresholdClaimAllowed: false,
		})
		expect(handoff.streamingMatrixEvidence).toMatchObject({
			status: 'representative-proof-present-owner-approval-required',
			ownerApprovalRequired: true,
			coveredActionKinds: ['passthrough', 'regenerate', 'add', 'drop'],
			missingActionKinds: ['error'],
			representativeProofCases: 5,
			streamingRegenerateParts: 4,
			coveredCaseNames: [
				'docprops-passthrough',
				'add-sheet-part',
				'calc-chain-drop',
				'macro-passthrough',
				'chart-sidecar-accounting',
			],
			publicNonStreamingCaseNames: ['unknown-part-error'],
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
			'encrypted-files',
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
			'encrypted-files',
			'streaming-scope',
		])
		expect(handoff.trustCompletenessBoundaryEvidence).toMatchObject({
			status: 'boundary-pinned-owner-scope',
			validationCommand: 'bun test packages/sdk/src/release-trust-matrix.test.ts',
			doesNotCloseGates: ['product', 'performance', 'release'],
		})
		expect(
			handoff.trustCompletenessBoundaryEvidence.outOfScopeClasses.map((item) => item.name),
		).toContain('Reader/writer performance and benchmark tuning')
		expect(JSON.stringify(handoff.fixturePolicy)).toContain('package-action-fixture-scan')
		expect(handoff.fixturePolicyEvidence.safeOpen.externalCandidateEvidence).toEqual([
			expect.objectContaining({
				candidateId: 'excelforge-book1-unknown-part',
				status: 'vendored-public-fixture',
				gateEffect: 'satisfies-unknown-part-only',
				riskFamily: 'preservedOther',
				recommendedMode: 'metadata-only',
			}),
		])
		expect(handoff.fixturePolicyEvidence.packageAction.externalCandidateEvidence).toEqual([
			expect.objectContaining({
				candidateId: 'excelforge-book1-unknown-part-mutation',
				status: 'vendored-public-fixture',
				gateEffect: 'satisfies-unknown-part-only',
				postWriteAuditsPassed: false,
				unknownPartPath: 'docMetadata/LabelInfo.xml',
			}),
		])
		expect(handoff.fixtureAcquisitionPlan.tasks[0]).toMatchObject({
			caseName: 'signed-package',
			killCriterion: expect.stringContaining('certificate meaning'),
		})
		expect(handoff.fixturePolicyEvidence.safeOpen.signatureOrUnknownMatches).toBe(1)
		expect(JSON.stringify(handoff.generatedFixtureDecisionEvidence)).toContain(
			'generated-malformed-package',
		)
		expect(JSON.stringify(handoff.generatedFixtureDecisionEvidence)).toContain(
			'none are accepted public malformed-package replacements',
		)
		expect(JSON.stringify(handoff.performancePolicy)).toContain('safe-open-proof.ts --repeat 10')
		expect(JSON.stringify(handoff.safeOpenLatencyValidationEvidence)).toContain(
			'tracked-clean release environment',
		)
		expect(JSON.stringify(handoff.streamingMatrixEvidence)).toContain(
			'"missingActionKinds":["error"]',
		)
		expect(JSON.stringify(handoff.correctnessPolicy)).toContain('signature preservation')
		expect(JSON.stringify(handoff.correctnessBoundaryEvidence)).toContain('safe-open-proof/activex')
		expect(handoff.releasePackageabilityEvidence).toMatchObject({
			status: 'local-tarball-smokes-present-publication-policy-required',
			sdkSmokeCommand: 'bun run release:sdk:smoke',
			appSmokeCommand: 'bun run release:apps:smoke',
			rcGateCommand: 'bun run release:rc:gate',
		})
		expect(JSON.stringify(handoff.releasePackageabilityEvidence)).toContain(
			'full MCP protocol compatibility',
		)
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
		expect(handoff.researchHygieneDecisionPacket).toMatchObject({
			ownerLoops: ['product', 'release'],
			status: 'claim-downgrade-do-not-promote',
			releaseGate: 'blocked-by-publication-policy',
			headlineClaimsAllowed: false,
			ownerApprovalRequired: true,
			claim: 'research-surface-hygiene',
			workBlockDisposition: 'claim-downgrade-do-not-promote',
			dirtyInventoryCommand:
				'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
		})
		expect(handoff.researchHygieneDecisionPacket.inventorySnapshot).toMatchObject({
			status: 'inventory-collected',
			decision: expect.stringMatching(
				/^(owner-classification-required|no-unclassified-paths-currently-visible)$/,
			),
		})
		expect(handoff.researchHygieneDecisionPacket.inventorySnapshot.dirtyPathCount).toBe(
			handoff.researchHygieneDecisionPacket.inventorySnapshot.unclassifiedEntries.length,
		)
		expect(
			handoff.researchHygieneDecisionPacket.classificationBuckets.map((item) => item.bucket),
		).toEqual(['accepted-evidence', 'active-owner-blocker', 'archive-only'])
		expect(handoff.researchHygieneDecisionPacket.validationCommands).toEqual([
			'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
			'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
			'bun test fixtures/benchmarks/release-proof-index.test.ts',
		])
		expect(handoff.researchHygieneDecisionPacket.boundary).toContain(
			'Compact research hygiene decision packet',
		)
		expect(handoff.nextOwnerActions[0]).toMatchObject({
			requirementId: 'edge-fixture-policy',
			acceptanceEvidence: expect.stringContaining('accepts disclosed generated'),
			forbiddenShortcut: expect.stringContaining('Do not hide generated fixture provenance'),
		})
		expect(handoff.claimBlockerBoard.map((row) => `${row.artifact}/${row.ownerLoop}`)).toEqual([
			'safe-open-proof/performance',
			'safe-open-proof/product',
			'safe-open-proof/release',
			'package-action-proof/correctness',
			'package-action-proof/performance',
			'package-action-proof/product',
			'package-action-proof/release',
		])
		expect(
			handoff.claimBlockerBoard.find(
				(row) => row.artifact === 'safe-open-proof' && row.ownerLoop === 'release',
			),
		).toMatchObject({
			requirementIds: ['publication-boundary', 'compact-report-publication-policy'],
			nextStepKinds: ['publication-policy'],
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
		expect(handoff.qssLeapfrogReleaseMatrix.rows.map((row) => row.artifact)).toEqual([
			'safe-open-proof',
			'package-action-proof',
		])
		expect(handoff.qssLeapfrogReleaseMatrix.rows[0].acceptedEvidence).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ evidenceId: 'safe-open-proof-tests' }),
				expect.objectContaining({ evidenceId: 'release-proof-index-owner-handoff' }),
			]),
		)
		expect(handoff.qssLeapfrogReleaseMatrix.rows[1].weakClaimDisposition).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					weakClaim: 'complete streaming package-action parity',
					disposition: 'blocker',
				}),
				expect.objectContaining({
					weakClaim: 'proof bundle as provenance or attestation',
					disposition: 'kill',
				}),
			]),
		)
		expect(handoff.qssLeapfrogReleaseMatrix.archivedResearchNotes).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					name: 'formula-language-service-primitives',
					status: 'archived-research-note',
				}),
				expect.objectContaining({
					name: 'practical-latency-contracts',
					status: 'archived-research-note',
				}),
				expect.objectContaining({
					name: 'research-surface-hygiene',
					status: 'archived-research-note',
				}),
			]),
		)
		expect(handoff.releaseDecisionBoard.rows).toHaveLength(2)
		expect(handoff.releaseDecisionBoard.rows[0]).toMatchObject({
			artifact: 'safe-open-proof',
			claimWordingAllowedToday: 'safe unknown workbook opening',
			headlineClaimAllowed: false,
			implementationSurfacePromotionAllowed: false,
			acceptedEvidence: expect.arrayContaining([
				expect.objectContaining({ evidenceId: 'release-proof-index-owner-handoff' }),
			]),
			ownerDecisionArtifacts: expect.arrayContaining([
				expect.objectContaining({
					artifactId: 'excel-behavior-compatibility-matrix',
				}),
				expect.objectContaining({
					artifactId: 'performance-claim-baseline-matrix',
				}),
			]),
			aPlusBlockingOwnerActions: expect.arrayContaining([
				expect.objectContaining({ requirementId: 'public-edge-fixtures' }),
				expect.objectContaining({ requirementId: 'release-latency-run' }),
			]),
		})
		expect(handoff.releaseDecisionBoard.rows[1]).toMatchObject({
			artifact: 'package-action-proof',
			claimWordingAllowedToday: 'auditable package-part mutation',
			ownerDecisionArtifacts: expect.arrayContaining([
				expect.objectContaining({
					artifactId: 'package-action-streaming-matrix-evidence',
				}),
			]),
			aPlusBlockingOwnerActions: expect.arrayContaining([
				expect.objectContaining({ requirementId: 'edge-fixture-policy' }),
				expect.objectContaining({ requirementId: 'provenance-boundary' }),
			]),
		})
		expect(handoff.releaseDecisionBoard.topClaimOwnerActionQueue).toHaveLength(9)
		for (const row of handoff.releaseDecisionBoard.topClaimOwnerActionQueue) {
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(Object)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction).toContain(row.validationCommand)
			expect(row.nextOwnerAction).toContain(row.acceptanceEvidence)
		}
		expect(handoff.releaseDecisionBoard.topClaimOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					ownerLoop: 'product',
					artifact: 'safe-open-proof',
					requirementId: 'public-edge-fixtures',
					workBlockDisposition: 'benchmark-corpus-blocker',
					validationCommand: 'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
				}),
				expect.objectContaining({
					ownerLoop: 'correctness',
					artifact: 'package-action-proof',
					requirementId: 'unsupported-feature-boundary',
					workBlockDisposition: 'implementation-ready-blocker',
					acceptanceEvidence: expect.stringContaining(
						'Correctness approves allowed/forbidden wording',
					),
				}),
				expect.objectContaining({
					ownerLoop: 'performance',
					artifact: 'package-action-proof',
					requirementId: 'streaming-matrix-boundary',
					workBlockDisposition: 'benchmark-corpus-blocker',
					validationCommand:
						'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
				}),
				expect.objectContaining({
					ownerLoop: 'release',
					artifact: 'package-action-proof',
					requirementId: 'provenance-boundary',
					workBlockDisposition: 'implementation-ready-blocker',
					forbiddenShortcut: expect.stringContaining('signed provenance'),
				}),
			]),
		)
		expect(handoff.releaseDecisionBoard.doNotPromoteYet).toHaveLength(10)
		for (const item of handoff.releaseDecisionBoard.doNotPromoteYet) {
			const qssContrast = item.qssContrast.join('\n')
			expect([
				'implementation-ready-blocker',
				'benchmark-corpus-blocker',
				'claim-downgrade-do-not-promote',
			]).toContain(item.workBlockDisposition)
			expect(item.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(qssContrast).not.toContain(
				'QSS contrast is blocked until this diagnostic evidence changes a top-two release claim.',
			)
			expect(item.allowedWording).toContain('Do not promote')
			expect(item.allowedWording).not.toContain('owner planning or research evidence')
			expect(item.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.nextOwnerAction.length).toBeGreaterThan(0)
			expect(item.validationCommands).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.validationCommands.every((command) => command.length > 0)).toBe(true)
			expect(item.nextOwnerAction).not.toContain(
				'No owner action is release-blocking until this claim changes the top-two release gate.',
			)
		}
		expect(handoff.releaseDecisionBoard.doNotPromoteYet).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					name: 'research-surface-hygiene',
					nextOwnerAction: expect.stringContaining(
						'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
					),
				}),
				expect.objectContaining({
					name: 'practical-latency-contracts',
					nextOwnerAction: expect.stringContaining(
						'--input-preset public-tracked --contract all --repeat 3 --warmup 1 --json',
					),
				}),
			]),
		)
		expect(handoff.releaseDecisionBoard.releaseWordingDecisionSummary).toMatchObject({
			status: 'headline-claims-blocked-local-wording-only',
			headlineClaimsAllowed: false,
			localAllowedClaimNames: ['safe-open-proof', 'package-action-proof'],
			doNotPromoteClaimNames: handoff.releaseDecisionBoard.doNotPromoteYet.map((item) => item.name),
		})
		expect(
			handoff.releaseDecisionBoard.releaseWordingDecisionSummary.forbiddenWordingByClaim[
				'formula-language-service-primitives'
			],
		).toEqual(
			handoff.releaseDecisionBoard.doNotPromoteYet.find(
				(item) => item.name === 'formula-language-service-primitives',
			)?.forbiddenWording,
		)
		expect(handoff.releaseDecisionBoard.claimDecisionContractCoverage).toMatchObject({
			status: 'all-release-claim-decisions-self-contained',
			decisionCount: 12,
			topClaimDecisionCount: 2,
			doNotPromoteDecisionCount: 10,
			missingEvidenceWeHaveKeys: [],
			missingEvidenceMissingKeys: [],
			missingQssContrastKeys: [],
			missingAllowedWordingKeys: [],
			missingForbiddenWordingKeys: [],
			missingNextOwnerActionKeys: [],
		})
		expect(handoff.releaseDecisionBoard.doNotPromoteDispositionSummary).toMatchObject({
			implementationReadyBlockerNames: [
				'formula-language-service-primitives',
				'token-bounded-agent-view',
				'retained-viewport-patch-history',
				'release-proof-bundle',
				'property-journal-laws',
				'agent-workflow-observability',
			],
			benchmarkCorpusBlockerNames: [
				'formula-oracle-routing',
				'columnar-scan-sidecars',
				'practical-latency-contracts',
			],
			claimDowngradeDoNotPromoteNames: ['research-surface-hygiene'],
		})
		expect(handoff.releaseDecisionBoard.blockedOwnerActionQueue).toHaveLength(14)
		for (const row of handoff.releaseDecisionBoard.blockedOwnerActionQueue) {
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
		}
		expect(handoff.releaseDecisionBoard.blockedOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					ownerLoop: 'correctness',
					name: 'formula-oracle-routing',
					workBlockDisposition: 'benchmark-corpus-blocker',
					validationCommands: expect.arrayContaining([
						expect.stringContaining('formula-corpus-correctness.ts'),
					]),
				}),
				expect.objectContaining({
					ownerLoop: 'performance',
					name: 'columnar-scan-sidecars',
					workBlockDisposition: 'benchmark-corpus-blocker',
					validationCommands: expect.arrayContaining([
						expect.stringContaining('columnar-sidecar.ts'),
					]),
				}),
				expect.objectContaining({
					ownerLoop: 'performance',
					name: 'practical-latency-contracts',
					workBlockDisposition: 'benchmark-corpus-blocker',
					validationCommands: expect.arrayContaining([
						expect.stringContaining(
							'--input-preset public-tracked --contract all --repeat 3 --warmup 1 --json',
						),
					]),
				}),
				expect.objectContaining({
					ownerLoop: 'release',
					name: 'research-surface-hygiene',
					workBlockDisposition: 'claim-downgrade-do-not-promote',
					validationCommands: [
						'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
						'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
						'bun test fixtures/benchmarks/release-proof-index.test.ts',
					],
				}),
			]),
		)
		expect(handoff.releaseDecisionBoard.benchmarkCorpusOwnerActionQueue).toHaveLength(7)
		for (const row of handoff.releaseDecisionBoard.benchmarkCorpusOwnerActionQueue) {
			expect(row.ownerFiles).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.commandsToRun).toEqual(row.validationCommands)
			expect(row.failureEvidence).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.acceptanceCriteria.length).toBeGreaterThan(0)
		}
		expect(handoff.releaseDecisionBoard.benchmarkCorpusOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					sourceQueue: 'top-claim-owner-action',
					ownerLoop: 'product',
					name: 'safe-open-proof',
					requirementId: 'public-edge-fixtures',
					ownerFiles: ['fixtures/benchmarks/safe-open-fixture-scan.ts'],
					validationCommands: ['bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json'],
				}),
				expect.objectContaining({
					sourceQueue: 'top-claim-owner-action',
					ownerLoop: 'performance',
					name: 'package-action-proof',
					requirementId: 'streaming-matrix-boundary',
					ownerFiles: ['fixtures/benchmarks/package-action-proof.ts'],
					validationCommands: [
						'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
					],
				}),
				expect.objectContaining({
					sourceQueue: 'blocked-owner-action',
					ownerLoop: 'correctness',
					name: 'formula-oracle-routing',
					ownerFiles: expect.arrayContaining([
						'fixtures/benchmarks/formula-corpus-correctness.ts',
						'fixtures/benchmarks/formula-corpus-correctness.test.ts',
					]),
					validationCommands: expect.arrayContaining([
						expect.stringContaining('formula-corpus-correctness.ts'),
					]),
				}),
				expect.objectContaining({
					sourceQueue: 'blocked-owner-action',
					ownerLoop: 'performance',
					name: 'columnar-scan-sidecars',
					ownerFiles: expect.arrayContaining(['fixtures/benchmarks/columnar-sidecar.ts']),
					validationCommands: expect.arrayContaining([
						expect.stringContaining('columnar-sidecar.ts'),
					]),
				}),
			]),
		)
		expect(handoff.releaseDecisionBoard.implementationReadyOwnerActionQueue).toHaveLength(14)
		for (const row of handoff.releaseDecisionBoard.implementationReadyOwnerActionQueue) {
			expect(row.ownerFiles).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.commandsToRun).toEqual(row.validationCommands)
			expect(row.failureEvidence).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.acceptanceCriteria.length).toBeGreaterThan(0)
		}
		expect(handoff.releaseDecisionBoard.implementationReadyOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					sourceQueue: 'top-claim-owner-action',
					ownerLoop: 'correctness',
					name: 'package-action-proof',
					requirementId: 'unsupported-feature-boundary',
					ownerFiles: ['fixtures/benchmarks/package-action-proof.ts'],
					validationCommands: [
						'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
					],
				}),
				expect.objectContaining({
					sourceQueue: 'top-claim-owner-action',
					ownerLoop: 'release',
					name: 'safe-open-proof',
					requirementId: 'compact-report-publication-policy',
					ownerFiles: ['fixtures/benchmarks/safe-open-proof.ts'],
					validationCommands: [
						'bun run fixtures/benchmarks/safe-open-proof.ts --no-timings --compact-json',
					],
				}),
				expect.objectContaining({
					sourceQueue: 'blocked-owner-action',
					ownerLoop: 'product',
					name: 'agent-workflow-observability',
					ownerFiles: expect.arrayContaining(['packages/sdk/src/agent-workflow.test.ts']),
					validationCommands: expect.arrayContaining([
						expect.stringContaining('packages/sdk/src/agent-workflow.test.ts'),
					]),
				}),
				expect.objectContaining({
					sourceQueue: 'blocked-owner-action',
					ownerLoop: 'release',
					name: 'release-proof-bundle',
					ownerFiles: expect.arrayContaining(['scripts/release-rc-gate.ts']),
					validationCommands: expect.arrayContaining(['bun run release:rc:gate']),
				}),
			]),
		)
		expect(handoff.releaseDecisionBoard.claimDowngradeOwnerActionQueue).toEqual([
			expect.objectContaining({
				sourceQueue: 'blocked-owner-action',
				ownerLoop: 'product',
				name: 'research-surface-hygiene',
				workBlockDisposition: 'claim-downgrade-do-not-promote',
				ownerFiles: expect.arrayContaining(['research/', 'tmp/ascend-loop-manager/']),
				validationCommands: [
					'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
					'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
					'bun test fixtures/benchmarks/release-proof-index.test.ts',
				],
				commandsToRun: [
					'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
					'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
					'bun test fixtures/benchmarks/release-proof-index.test.ts',
				],
				failureEvidence: expect.arrayContaining([
					expect.stringContaining('Classify current research files'),
				]),
			}),
			expect.objectContaining({
				sourceQueue: 'blocked-owner-action',
				ownerLoop: 'release',
				name: 'research-surface-hygiene',
				workBlockDisposition: 'claim-downgrade-do-not-promote',
				ownerFiles: expect.arrayContaining(['research/', 'tmp/ascend-loop-manager/']),
				validationCommands: [
					'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
					'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
					'bun test fixtures/benchmarks/release-proof-index.test.ts',
				],
				acceptanceCriteria: expect.stringContaining('classifies each untriaged path'),
			}),
		])
		expect(handoff.releaseDecisionBoard.ownerActionQueueCoverage).toMatchObject({
			status: 'all-owner-actions-covered-by-disposition-queues',
			sourceTopClaimActionCount: 9,
			sourceBlockedActionCount: 14,
			benchmarkCorpusActionCount: 7,
			implementationReadyActionCount: 14,
			claimDowngradeActionCount: 2,
			coveredActionCount: 23,
			uncoveredTopClaimActionKeys: [],
			uncoveredBlockedActionKeys: [],
		})
		expect(handoff.releaseDecisionBoard.ownerActionExecutionContractCoverage).toMatchObject({
			status: 'all-disposition-owner-actions-have-execution-contract',
			actionCount: 23,
			benchmarkCorpusActionCount: 7,
			implementationReadyActionCount: 14,
			claimDowngradeActionCount: 2,
			missingOwnerFileActionKeys: [],
			missingCommandActionKeys: [],
			missingFailureEvidenceActionKeys: [],
			missingAcceptanceCriteriaActionKeys: [],
		})
		const releaseDecisionCoverage = new Set([
			...handoff.releaseDecisionBoard.rows.map((row) => row.artifact),
			...handoff.releaseDecisionBoard.doNotPromoteYet.map((item) => item.name),
		])
		expect(handoff.claimDecisionCoverage).toMatchObject({
			status: 'all-handoff-claims-covered-by-release-decision-board',
			portfolioClaimCount: 11,
			deferredClaimCount: 7,
			excludedEvidenceCount: 1,
			uncoveredPortfolioClaimNames: [],
			uncoveredDeferredClaimNames: [],
			uncoveredExcludedEvidenceNames: [],
			boundary: expect.stringContaining('Every owner-handoff claim'),
		})
		expect(handoff.claimDecisionCoverage.topClaimNames).toEqual(
			handoff.releaseDecisionBoard.rows.map((row) => row.artifact),
		)
		expect(handoff.claimDecisionCoverage.doNotPromoteNames).toEqual(
			handoff.releaseDecisionBoard.doNotPromoteYet.map((item) => item.name),
		)
		expect(
			handoff.claimPortfolio
				.map((claim) => claim.name)
				.filter((name) => !releaseDecisionCoverage.has(name)),
		).toEqual([])
		expect(
			handoff.deferredClaims
				.map((claim) => claim.name)
				.filter((name) => !releaseDecisionCoverage.has(name)),
		).toEqual([])
		expect(
			handoff.excludedEvidence
				.map((evidence) => evidence.name)
				.filter((name) => !releaseDecisionCoverage.has(name)),
		).toEqual([])
		for (const claim of handoff.claimPortfolio.filter(
			(entry) => entry.handoffDecision === 'top-implementation-handoff',
		)) {
			const decision = handoff.releaseDecisionBoard.rows.find((row) => row.artifact === claim.name)
			if (!decision) {
				throw new Error(`Missing release decision row for ${claim.name}`)
			}
			expect(decision.headlineClaimAllowed).toBe(false)
			expect(decision.implementationSurfacePromotionAllowed).toBe(false)
			expect(decision.evidenceWeHave.length).toBeGreaterThan(0)
			expect(decision.evidenceWeHave.every((item) => item.evidenceId.length > 0)).toBe(true)
			expect(decision.evidenceMissing.length).toBeGreaterThan(0)
			expect(decision.qssContrast.length).toBeGreaterThan(0)
			expect(decision.allowedWording.length).toBeGreaterThan(0)
			expect(decision.forbiddenWording.length).toBeGreaterThan(0)
			expect(decision.nextOwnerActions.length).toBeGreaterThan(0)
			expect(decision.nextOwnerActions.every((action) => action.requirementId.length > 0)).toBe(
				true,
			)
		}
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
		expect(handoff.claimPortfolio).toHaveLength(11)
		expect(handoff.claimPortfolio[0]).toMatchObject({
			name: 'safe-open-proof',
			handoffDecision: 'top-implementation-handoff',
			evidenceNeeded: {
				competitorContrast: expect.stringContaining('Microsoft Protected View'),
			},
		})
		expect(handoff.claimPortfolio[7]).toMatchObject({
			name: 'property-journal-laws',
			status: 'speculative-do-not-promote',
			killCriterion: expect.stringContaining('Do not promote broad inverse-law claims'),
		})
		expect(handoff.claimPortfolio[10]).toMatchObject({
			name: 'research-surface-hygiene',
			status: 'speculative-do-not-promote',
			handoffDecision: 'do-not-promote-yet',
		})
		expect(handoff.deferredClaims.map((entry) => entry.status)).toContain('do-not-promote-yet')
		expect(handoff.excludedEvidence.map((entry) => entry.name)).toEqual([
			'practical-latency-contracts',
		])
		expect(JSON.stringify(handoff)).not.toContain('"artifacts"')
		expect(JSON.stringify(handoff)).toContain('"claimBlockerBoard"')
		expect(JSON.stringify(handoff)).toContain('"claimDecisionCoverage"')
		expect(JSON.stringify(handoff)).toContain('"nextOwnerActions"')
	})

	test('emits a compact release decision board JSON mode', async () => {
		const proc = Bun.spawn([Bun.argv[0], runnerPath, '--no-timings', '--release-decision-json'], {
			cwd: process.cwd(),
			stderr: 'pipe',
			stdout: 'pipe',
		})
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])

		expect(exitCode, stderr).toBe(0)
		expect(stderr.trim()).toBe('')
		const board = JSON.parse(stdout) as {
			readonly status?: string
			readonly releaseGate?: string
			readonly headlineClaimsAllowed?: boolean
			readonly implementationSurfacePromotionAllowed?: boolean
			readonly missingRequirementCount?: number
			readonly rows?: readonly {
				readonly artifact?: string
				readonly evidenceWeHave?: readonly unknown[]
				readonly evidenceMissing?: readonly string[]
				readonly qssContrast?: readonly string[]
				readonly allowedWording?: string
				readonly forbiddenWording?: readonly string[]
				readonly nextOwnerActions?: readonly unknown[]
				readonly ownerDecisionArtifacts?: readonly unknown[]
				readonly headlineClaimAllowed?: boolean
				readonly implementationSurfacePromotionAllowed?: boolean
				readonly aPlusBlockingOwnerActions?: readonly unknown[]
				readonly claimsWeMustNotMake?: readonly string[]
			}[]
			readonly topClaimOwnerActionQueue?: readonly {
				readonly artifact?: string
				readonly ownerLoop?: string
				readonly requirementId?: string
				readonly workBlockDisposition?: string
				readonly nextStepKind?: string
				readonly evidenceWeHave?: readonly unknown[]
				readonly evidenceMissing?: readonly string[]
				readonly qssContrast?: readonly string[]
				readonly validationCommand?: string
				readonly acceptanceEvidence?: string
				readonly forbiddenShortcut?: string
				readonly allowedWording?: string
				readonly forbiddenWording?: readonly string[]
				readonly nextOwnerAction?: string
			}[]
			readonly doNotPromoteYet?: readonly {
				readonly name?: string
				readonly status?: string
				readonly workBlockDisposition?: string
				readonly evidenceWeHave?: readonly string[]
				readonly evidenceMissing?: readonly string[]
				readonly qssContrast?: readonly string[]
				readonly allowedWording?: string
				readonly forbiddenWording?: readonly string[]
				readonly nextOwnerAction?: string
				readonly validationCommands?: readonly string[]
			}[]
			readonly doNotPromoteDispositionSummary?: {
				readonly implementationReadyBlockerNames?: readonly string[]
				readonly benchmarkCorpusBlockerNames?: readonly string[]
				readonly claimDowngradeDoNotPromoteNames?: readonly string[]
			}
			readonly releaseWordingDecisionSummary?: {
				readonly status?: string
				readonly headlineClaimsAllowed?: boolean
				readonly localAllowedClaimNames?: readonly string[]
				readonly doNotPromoteClaimNames?: readonly string[]
				readonly localAllowedWordingByClaim?: Record<string, string>
				readonly doNotPromoteAllowedWordingByClaim?: Record<string, string>
				readonly forbiddenWordingByClaim?: Record<string, readonly string[]>
			}
			readonly claimDecisionContractCoverage?: {
				readonly status?: string
				readonly decisionCount?: number
				readonly topClaimDecisionCount?: number
				readonly doNotPromoteDecisionCount?: number
				readonly missingEvidenceWeHaveKeys?: readonly string[]
				readonly missingEvidenceMissingKeys?: readonly string[]
				readonly missingQssContrastKeys?: readonly string[]
				readonly missingAllowedWordingKeys?: readonly string[]
				readonly missingForbiddenWordingKeys?: readonly string[]
				readonly missingNextOwnerActionKeys?: readonly string[]
			}
			readonly blockedOwnerActionQueue?: readonly {
				readonly name?: string
				readonly ownerLoop?: string
				readonly workBlockDisposition?: string
				readonly ownerFiles?: readonly string[]
				readonly validationCommands?: readonly string[]
				readonly commandsToRun?: readonly string[]
				readonly failureEvidence?: readonly string[]
				readonly acceptanceCriteria?: string
				readonly evidenceWeHave?: readonly string[]
				readonly evidenceMissing?: readonly string[]
				readonly qssContrast?: readonly string[]
				readonly allowedWording?: string
				readonly forbiddenWording?: readonly string[]
				readonly nextOwnerAction?: string
			}[]
			readonly benchmarkCorpusOwnerActionQueue?: readonly {
				readonly sourceQueue?: string
				readonly name?: string
				readonly ownerLoop?: string
				readonly requirementId?: string
				readonly workBlockDisposition?: string
				readonly ownerFiles?: readonly string[]
				readonly validationCommands?: readonly string[]
				readonly commandsToRun?: readonly string[]
				readonly failureEvidence?: readonly string[]
				readonly acceptanceCriteria?: string
				readonly evidenceWeHave?: readonly string[]
				readonly evidenceMissing?: readonly string[]
				readonly qssContrast?: readonly string[]
				readonly allowedWording?: string
				readonly forbiddenWording?: readonly string[]
				readonly nextOwnerAction?: string
			}[]
			readonly implementationReadyOwnerActionQueue?: readonly {
				readonly sourceQueue?: string
				readonly name?: string
				readonly ownerLoop?: string
				readonly requirementId?: string
				readonly workBlockDisposition?: string
				readonly ownerFiles?: readonly string[]
				readonly validationCommands?: readonly string[]
				readonly commandsToRun?: readonly string[]
				readonly failureEvidence?: readonly string[]
				readonly acceptanceCriteria?: string
				readonly evidenceWeHave?: readonly string[]
				readonly evidenceMissing?: readonly string[]
				readonly qssContrast?: readonly string[]
				readonly allowedWording?: string
				readonly forbiddenWording?: readonly string[]
				readonly nextOwnerAction?: string
			}[]
			readonly claimDowngradeOwnerActionQueue?: readonly {
				readonly sourceQueue?: string
				readonly name?: string
				readonly ownerLoop?: string
				readonly workBlockDisposition?: string
				readonly validationCommands?: readonly string[]
				readonly evidenceWeHave?: readonly string[]
				readonly evidenceMissing?: readonly string[]
				readonly qssContrast?: readonly string[]
				readonly allowedWording?: string
				readonly forbiddenWording?: readonly string[]
				readonly nextOwnerAction?: string
			}[]
			readonly ownerActionQueueCoverage?: {
				readonly status?: string
				readonly sourceTopClaimActionCount?: number
				readonly sourceBlockedActionCount?: number
				readonly benchmarkCorpusActionCount?: number
				readonly implementationReadyActionCount?: number
				readonly claimDowngradeActionCount?: number
				readonly coveredActionCount?: number
				readonly uncoveredTopClaimActionKeys?: readonly string[]
				readonly uncoveredBlockedActionKeys?: readonly string[]
			}
			readonly ownerActionExecutionContractCoverage?: {
				readonly status?: string
				readonly actionCount?: number
				readonly benchmarkCorpusActionCount?: number
				readonly implementationReadyActionCount?: number
				readonly claimDowngradeActionCount?: number
				readonly missingOwnerFileActionKeys?: readonly string[]
				readonly missingCommandActionKeys?: readonly string[]
				readonly missingFailureEvidenceActionKeys?: readonly string[]
				readonly missingAcceptanceCriteriaActionKeys?: readonly string[]
			}
		}
		expect(board.status).toBe('top-two-only')
		expect(board.releaseGate).toBe('blocked-by-publication-policy')
		expect(board.headlineClaimsAllowed).toBe(false)
		expect(board.implementationSurfacePromotionAllowed).toBe(false)
		expect(board.missingRequirementCount).toBe(9)
		expect(board.rows?.map((row) => row.artifact)).toEqual([
			'safe-open-proof',
			'package-action-proof',
		])
		expect(board.rows?.every((row) => row.headlineClaimAllowed === false)).toBe(true)
		expect(board.rows?.every((row) => row.implementationSurfacePromotionAllowed === false)).toBe(
			true,
		)
		expect(board.rows?.every((row) => (row.aPlusBlockingOwnerActions?.length ?? 0) > 0)).toBe(true)
		for (const row of board.rows ?? []) {
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(Object)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.allowedWording).toEqual(expect.any(String))
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerActions).toEqual(expect.arrayContaining([expect.any(Object)]))
			expect(row.ownerDecisionArtifacts).toEqual(expect.arrayContaining([expect.any(Object)]))
			expect(row.claimsWeMustNotMake).toEqual(expect.arrayContaining([expect.any(String)]))
		}
		expect(board.rows?.[0]).toMatchObject({
			artifact: 'safe-open-proof',
			allowedWording: expect.stringContaining('pre-hydration package-feature routing'),
			evidenceMissing: expect.arrayContaining([
				expect.stringContaining('public-edge-fixtures'),
				expect.stringContaining('release-latency-run'),
			]),
			qssContrast: expect.arrayContaining([expect.stringContaining('QSS likely does well')]),
			forbiddenWording: expect.arrayContaining([
				expect.stringContaining('Microsoft Protected View equivalence'),
			]),
		})
		expect(board.rows?.[1]).toMatchObject({
			artifact: 'package-action-proof',
			allowedWording: expect.stringContaining('local per-part package action accounting'),
			evidenceMissing: expect.arrayContaining([
				expect.stringContaining('edge-fixture-policy'),
				expect.stringContaining('streaming-matrix-boundary'),
			]),
			qssContrast: expect.arrayContaining([expect.stringContaining('QSS likely does well')]),
			forbiddenWording: expect.arrayContaining([
				expect.stringContaining('Signed provenance'),
				expect.stringContaining('Full streaming parity'),
			]),
		})
		expect(board.releaseWordingDecisionSummary).toMatchObject({
			status: 'headline-claims-blocked-local-wording-only',
			headlineClaimsAllowed: false,
			localAllowedClaimNames: ['safe-open-proof', 'package-action-proof'],
			doNotPromoteClaimNames: board.doNotPromoteYet?.map((item) => item.name),
		})
		expect(
			board.releaseWordingDecisionSummary?.localAllowedWordingByClaim?.['safe-open-proof'],
		).toContain('pre-hydration package-feature routing')
		expect(
			board.releaseWordingDecisionSummary?.forbiddenWordingByClaim?.['package-action-proof'],
		).toEqual(board.rows?.[1].forbiddenWording)
		expect(
			board.releaseWordingDecisionSummary?.doNotPromoteAllowedWordingByClaim?.[
				'research-surface-hygiene'
			],
		).toContain('Do not promote research-surface-hygiene')
		expect(board.claimDecisionContractCoverage).toMatchObject({
			status: 'all-release-claim-decisions-self-contained',
			decisionCount: 12,
			topClaimDecisionCount: 2,
			doNotPromoteDecisionCount: 10,
			missingEvidenceWeHaveKeys: [],
			missingEvidenceMissingKeys: [],
			missingQssContrastKeys: [],
			missingAllowedWordingKeys: [],
			missingForbiddenWordingKeys: [],
			missingNextOwnerActionKeys: [],
		})
		expect(board.topClaimOwnerActionQueue).toHaveLength(9)
		expect(
			board.topClaimOwnerActionQueue?.map(
				(row) =>
					`${row.ownerLoop}:${row.artifact}:${row.requirementId}:${row.nextStepKind}:${row.workBlockDisposition}`,
			),
		).toEqual([
			'product:safe-open-proof:public-edge-fixtures:owner-decision-or-fixture-replacement:benchmark-corpus-blocker',
			'performance:safe-open-proof:release-latency-run:validation-run:benchmark-corpus-blocker',
			'release:safe-open-proof:publication-boundary:publication-policy:implementation-ready-blocker',
			'release:safe-open-proof:compact-report-publication-policy:publication-policy:implementation-ready-blocker',
			'product:package-action-proof:edge-fixture-policy:owner-decision-or-fixture-replacement:benchmark-corpus-blocker',
			'correctness:package-action-proof:unsupported-feature-boundary:owner-boundary-approval:implementation-ready-blocker',
			'performance:package-action-proof:streaming-matrix-boundary:owner-decision-or-harness-expansion:benchmark-corpus-blocker',
			'release:package-action-proof:provenance-boundary:publication-policy:implementation-ready-blocker',
			'release:package-action-proof:compact-report-publication-policy:publication-policy:implementation-ready-blocker',
		])
		for (const row of board.topClaimOwnerActionQueue ?? []) {
			expect(['implementation-ready-blocker', 'benchmark-corpus-blocker']).toContain(
				row.workBlockDisposition,
			)
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(Object)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.validationCommand).toEqual(expect.any(String))
			expect(row.acceptanceEvidence).toEqual(expect.any(String))
			expect(row.forbiddenShortcut).toEqual(expect.any(String))
			expect(row.allowedWording).toEqual(expect.any(String))
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction).toEqual(expect.any(String))
			expect(row.nextOwnerAction).toContain(String(row.validationCommand))
			expect(row.nextOwnerAction).toContain(String(row.acceptanceEvidence))
		}
		expect(board.topClaimOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					ownerLoop: 'correctness',
					artifact: 'package-action-proof',
					requirementId: 'unsupported-feature-boundary',
					workBlockDisposition: 'implementation-ready-blocker',
					acceptanceEvidence: expect.stringContaining(
						'Correctness approves allowed/forbidden wording',
					),
				}),
				expect.objectContaining({
					ownerLoop: 'performance',
					artifact: 'safe-open-proof',
					requirementId: 'release-latency-run',
					workBlockDisposition: 'benchmark-corpus-blocker',
					validationCommand:
						'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
				}),
			]),
		)
		expect(stdout).toContain('Excel-ground-truth formula/cached-result fixtures')
		expect(stdout).toContain('bounded chart series-source fixtures')
		expect(stdout).toContain('public external-link/query-table refresh metadata fixtures')
		expect(stdout).toContain('full chart editing support')
		expect(stdout).toContain('Classify current research files')
		expect(board.doNotPromoteYet?.map((item) => item.name)).toEqual([
			'formula-language-service-primitives',
			'token-bounded-agent-view',
			'retained-viewport-patch-history',
			'release-proof-bundle',
			'formula-oracle-routing',
			'property-journal-laws',
			'columnar-scan-sidecars',
			'agent-workflow-observability',
			'research-surface-hygiene',
			'practical-latency-contracts',
		])
		expect(
			board.doNotPromoteYet?.map((item) => `${item.name}:${item.workBlockDisposition}`),
		).toEqual([
			'formula-language-service-primitives:implementation-ready-blocker',
			'token-bounded-agent-view:implementation-ready-blocker',
			'retained-viewport-patch-history:implementation-ready-blocker',
			'release-proof-bundle:implementation-ready-blocker',
			'formula-oracle-routing:benchmark-corpus-blocker',
			'property-journal-laws:implementation-ready-blocker',
			'columnar-scan-sidecars:benchmark-corpus-blocker',
			'agent-workflow-observability:implementation-ready-blocker',
			'research-surface-hygiene:claim-downgrade-do-not-promote',
			'practical-latency-contracts:benchmark-corpus-blocker',
		])
		expect(board.doNotPromoteDispositionSummary).toMatchObject({
			implementationReadyBlockerNames: [
				'formula-language-service-primitives',
				'token-bounded-agent-view',
				'retained-viewport-patch-history',
				'release-proof-bundle',
				'property-journal-laws',
				'agent-workflow-observability',
			],
			benchmarkCorpusBlockerNames: [
				'formula-oracle-routing',
				'columnar-scan-sidecars',
				'practical-latency-contracts',
			],
			claimDowngradeDoNotPromoteNames: ['research-surface-hygiene'],
		})
		expect(board.blockedOwnerActionQueue).toHaveLength(14)
		expect(
			board.blockedOwnerActionQueue?.map(
				(row) => `${row.ownerLoop}:${row.name}:${row.workBlockDisposition}`,
			),
		).toEqual([
			'product:formula-language-service-primitives:implementation-ready-blocker',
			'correctness:formula-language-service-primitives:implementation-ready-blocker',
			'product:token-bounded-agent-view:implementation-ready-blocker',
			'product:retained-viewport-patch-history:implementation-ready-blocker',
			'performance:retained-viewport-patch-history:implementation-ready-blocker',
			'product:release-proof-bundle:implementation-ready-blocker',
			'release:release-proof-bundle:implementation-ready-blocker',
			'correctness:formula-oracle-routing:benchmark-corpus-blocker',
			'correctness:property-journal-laws:implementation-ready-blocker',
			'performance:columnar-scan-sidecars:benchmark-corpus-blocker',
			'product:agent-workflow-observability:implementation-ready-blocker',
			'product:research-surface-hygiene:claim-downgrade-do-not-promote',
			'release:research-surface-hygiene:claim-downgrade-do-not-promote',
			'performance:practical-latency-contracts:benchmark-corpus-blocker',
		])
		for (const row of board.blockedOwnerActionQueue ?? []) {
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.allowedWording).toContain('Do not promote')
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction).toEqual(expect.any(String))
			expect(row.validationCommands).toEqual(expect.arrayContaining([expect.any(String)]))
		}
		expect(board.blockedOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					ownerLoop: 'performance',
					name: 'practical-latency-contracts',
					workBlockDisposition: 'benchmark-corpus-blocker',
					validationCommands: expect.arrayContaining([
						expect.stringContaining(
							'--input-preset public-tracked --contract all --repeat 3 --warmup 1 --json',
						),
					]),
				}),
			]),
		)
		expect(board.benchmarkCorpusOwnerActionQueue).toHaveLength(7)
		expect(
			board.benchmarkCorpusOwnerActionQueue?.map(
				(row) => `${row.sourceQueue}:${row.ownerLoop}:${row.name}:${row.requirementId ?? 'none'}`,
			),
		).toEqual([
			'top-claim-owner-action:product:safe-open-proof:public-edge-fixtures',
			'top-claim-owner-action:performance:safe-open-proof:release-latency-run',
			'top-claim-owner-action:product:package-action-proof:edge-fixture-policy',
			'top-claim-owner-action:performance:package-action-proof:streaming-matrix-boundary',
			'blocked-owner-action:correctness:formula-oracle-routing:none',
			'blocked-owner-action:performance:columnar-scan-sidecars:none',
			'blocked-owner-action:performance:practical-latency-contracts:none',
		])
		for (const row of board.benchmarkCorpusOwnerActionQueue ?? []) {
			expect(row.workBlockDisposition).toBe('benchmark-corpus-blocker')
			expect(row.ownerFiles).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.validationCommands).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.commandsToRun).toEqual(row.validationCommands)
			expect(row.failureEvidence).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.acceptanceCriteria).toEqual(expect.any(String))
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.allowedWording).toEqual(expect.any(String))
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction).toEqual(expect.any(String))
		}
		expect(board.benchmarkCorpusOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					sourceQueue: 'top-claim-owner-action',
					name: 'safe-open-proof',
					requirementId: 'release-latency-run',
					ownerFiles: ['fixtures/benchmarks/safe-open-proof.ts'],
					validationCommands: [
						'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
					],
					commandsToRun: [
						'bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json',
					],
					failureEvidence: expect.arrayContaining([expect.stringContaining('release-latency-run')]),
				}),
				expect.objectContaining({
					sourceQueue: 'blocked-owner-action',
					name: 'practical-latency-contracts',
					ownerFiles: [
						'fixtures/benchmarks/practical-latency-contracts.ts',
						'fixtures/benchmarks/practical-latency-contracts.test.ts',
					],
					validationCommands: expect.arrayContaining([
						expect.stringContaining(
							'--input-preset public-tracked --contract all --repeat 3 --warmup 1 --json',
						),
					]),
					acceptanceCriteria: expect.stringContaining('public-tracked'),
				}),
			]),
		)
		expect(board.implementationReadyOwnerActionQueue).toHaveLength(14)
		expect(
			board.implementationReadyOwnerActionQueue?.map(
				(row) => `${row.sourceQueue}:${row.ownerLoop}:${row.name}:${row.requirementId ?? 'none'}`,
			),
		).toEqual([
			'top-claim-owner-action:release:safe-open-proof:publication-boundary',
			'top-claim-owner-action:release:safe-open-proof:compact-report-publication-policy',
			'top-claim-owner-action:correctness:package-action-proof:unsupported-feature-boundary',
			'top-claim-owner-action:release:package-action-proof:provenance-boundary',
			'top-claim-owner-action:release:package-action-proof:compact-report-publication-policy',
			'blocked-owner-action:product:formula-language-service-primitives:none',
			'blocked-owner-action:correctness:formula-language-service-primitives:none',
			'blocked-owner-action:product:token-bounded-agent-view:none',
			'blocked-owner-action:product:retained-viewport-patch-history:none',
			'blocked-owner-action:performance:retained-viewport-patch-history:none',
			'blocked-owner-action:product:release-proof-bundle:none',
			'blocked-owner-action:release:release-proof-bundle:none',
			'blocked-owner-action:correctness:property-journal-laws:none',
			'blocked-owner-action:product:agent-workflow-observability:none',
		])
		for (const row of board.implementationReadyOwnerActionQueue ?? []) {
			expect(row.workBlockDisposition).toBe('implementation-ready-blocker')
			expect(row.ownerFiles).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.validationCommands).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.commandsToRun).toEqual(row.validationCommands)
			expect(row.failureEvidence).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.acceptanceCriteria).toEqual(expect.any(String))
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.allowedWording).toEqual(expect.any(String))
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction).toEqual(expect.any(String))
		}
		expect(board.implementationReadyOwnerActionQueue).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					sourceQueue: 'top-claim-owner-action',
					name: 'package-action-proof',
					requirementId: 'unsupported-feature-boundary',
					ownerFiles: ['fixtures/benchmarks/package-action-proof.ts'],
					validationCommands: [
						'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
					],
					commandsToRun: [
						'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
					],
					failureEvidence: expect.arrayContaining([
						expect.stringContaining('unsupported-feature-boundary'),
					]),
				}),
				expect.objectContaining({
					sourceQueue: 'blocked-owner-action',
					name: 'formula-language-service-primitives',
					ownerFiles: expect.arrayContaining([
						'fixtures/benchmarks/formula-assist-proof.ts',
						'packages/sdk/src/formula-edit.test.ts',
					]),
					validationCommands: expect.arrayContaining([
						expect.stringContaining('formula-assist-proof.ts'),
					]),
				}),
			]),
		)
		expect(board.claimDowngradeOwnerActionQueue).toEqual([
			expect.objectContaining({
				sourceQueue: 'blocked-owner-action',
				ownerLoop: 'product',
				name: 'research-surface-hygiene',
				workBlockDisposition: 'claim-downgrade-do-not-promote',
				ownerFiles: expect.arrayContaining(['research/', 'tmp/ascend-loop-manager/']),
				validationCommands: [
					'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
					'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
					'bun test fixtures/benchmarks/release-proof-index.test.ts',
				],
				commandsToRun: [
					'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
					'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
					'bun test fixtures/benchmarks/release-proof-index.test.ts',
				],
			}),
			expect.objectContaining({
				sourceQueue: 'blocked-owner-action',
				ownerLoop: 'release',
				name: 'research-surface-hygiene',
				workBlockDisposition: 'claim-downgrade-do-not-promote',
				ownerFiles: expect.arrayContaining(['research/', 'tmp/ascend-loop-manager/']),
				validationCommands: [
					'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
					'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
					'bun test fixtures/benchmarks/release-proof-index.test.ts',
				],
				acceptanceCriteria: expect.stringContaining('classifies each untriaged path'),
			}),
		])
		for (const row of board.claimDowngradeOwnerActionQueue ?? []) {
			expect(row.ownerFiles).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.commandsToRun).toEqual(row.validationCommands)
			expect(row.failureEvidence).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.acceptanceCriteria).toEqual(expect.any(String))
			expect(row.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.allowedWording).toEqual(expect.any(String))
			expect(row.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(row.nextOwnerAction).toEqual(expect.any(String))
		}
		expect(board.ownerActionQueueCoverage).toMatchObject({
			status: 'all-owner-actions-covered-by-disposition-queues',
			sourceTopClaimActionCount: 9,
			sourceBlockedActionCount: 14,
			benchmarkCorpusActionCount: 7,
			implementationReadyActionCount: 14,
			claimDowngradeActionCount: 2,
			coveredActionCount: 23,
			uncoveredTopClaimActionKeys: [],
			uncoveredBlockedActionKeys: [],
		})
		expect(board.ownerActionExecutionContractCoverage).toMatchObject({
			status: 'all-disposition-owner-actions-have-execution-contract',
			actionCount: 23,
			benchmarkCorpusActionCount: 7,
			implementationReadyActionCount: 14,
			claimDowngradeActionCount: 2,
			missingOwnerFileActionKeys: [],
			missingCommandActionKeys: [],
			missingFailureEvidenceActionKeys: [],
			missingAcceptanceCriteriaActionKeys: [],
		})
		expect(board.doNotPromoteYet?.every((item) => item.status === 'do-not-promote-yet')).toBe(true)
		for (const item of board.doNotPromoteYet ?? []) {
			expect([
				'implementation-ready-blocker',
				'benchmark-corpus-blocker',
				'claim-downgrade-do-not-promote',
			]).toContain(item.workBlockDisposition)
			expect(item.evidenceWeHave).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.evidenceMissing).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.qssContrast).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.qssContrast?.join('\n')).not.toContain(
				'QSS contrast is blocked until this diagnostic evidence changes a top-two release claim.',
			)
			expect(item.allowedWording).toContain('Do not promote')
			expect(item.allowedWording).not.toContain('owner planning or research evidence')
			expect(item.forbiddenWording).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.nextOwnerAction).toEqual(expect.any(String))
			expect(item.validationCommands).toEqual(expect.arrayContaining([expect.any(String)]))
			expect(item.nextOwnerAction).not.toContain(
				'No owner action is release-blocking until this claim changes the top-two release gate.',
			)
		}
		expect(board.doNotPromoteYet).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					name: 'research-surface-hygiene',
					nextOwnerAction: expect.stringContaining(
						'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
					),
				}),
				expect.objectContaining({
					name: 'practical-latency-contracts',
					nextOwnerAction: expect.stringContaining(
						'--input-preset public-tracked --contract all --repeat 3 --warmup 1 --json',
					),
				}),
			]),
		)
		expect(stdout).not.toContain('"claimBlockerBoard"')
		expect(stdout).not.toContain('"fixturePolicy"')
		expect(stdout).not.toContain('"deferredClaims"')
	})

	test('emits a compact fixture decision packet JSON mode', async () => {
		const proc = Bun.spawn([Bun.argv[0], runnerPath, '--no-timings', '--fixture-decision-json'], {
			cwd: process.cwd(),
			stderr: 'pipe',
			stdout: 'pipe',
		})
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])

		expect(exitCode, stderr).toBe(0)
		expect(stderr.trim()).toBe('')
		const packet = JSON.parse(stdout) as {
			readonly ownerLoop?: string
			readonly status?: string
			readonly releaseGate?: string
			readonly headlineClaimsAllowed?: boolean
			readonly ownerApprovalRequired?: boolean
			readonly publicReplacementGapsRemain?: boolean
			readonly trackedScans?: readonly {
				readonly artifact?: string
				readonly gateId?: string
				readonly publicReplacementGap?: boolean
				readonly generatedStructuralCases?: readonly string[]
			}[]
			readonly approvalChecklist?: readonly {
				readonly gateId?: string
				readonly ownerLoop?: string
			}[]
			readonly generatedCases?: readonly {
				readonly caseName?: string
				readonly forbiddenUse?: string
			}[]
			readonly validationCommands?: readonly string[]
			readonly forbiddenShortcuts?: readonly string[]
		}
		expect(packet).toMatchObject({
			ownerLoop: 'product',
			status: 'owner-decision-required',
			releaseGate: 'blocked-by-publication-policy',
			headlineClaimsAllowed: false,
			ownerApprovalRequired: true,
			publicReplacementGapsRemain: true,
		})
		expect(packet.trackedScans?.map((scan) => `${scan.artifact}/${scan.gateId}`)).toEqual([
			'safe-open-proof/public-edge-fixtures',
			'package-action-proof/edge-fixture-policy',
		])
		expect(packet.trackedScans?.every((scan) => scan.publicReplacementGap === true)).toBe(true)
		expect(packet.trackedScans?.[0].generatedStructuralCases).toEqual(['signed', 'malformed'])
		expect(packet.approvalChecklist?.map((item) => item.ownerLoop)).toEqual(['product', 'product'])
		expect(packet.approvalChecklist?.map((item) => item.gateId)).toEqual([
			'public-edge-fixtures',
			'edge-fixture-policy',
		])
		expect(packet.generatedCases?.map((item) => item.caseName)).toEqual([
			'signed',
			'malformed',
			'signature-invalidation-drop',
		])
		expect(packet.generatedCases?.[0].forbiddenUse).toContain('signature verification')
		expect(packet.validationCommands).toEqual([
			'bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json',
			'bun run fixtures/benchmarks/package-action-fixture-scan.ts --json',
			'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
		])
		expect(packet.forbiddenShortcuts?.join('\n')).toContain(
			'Do not hide generated fixture provenance',
		)
		expect(stdout).not.toContain('"claimBlockerBoard"')
		expect(stdout).not.toContain('"deferredClaims"')
		expect(stdout).not.toContain('"qssLeapfrogReleaseMatrix"')
	})

	test('emits a compact correctness boundary packet JSON mode', async () => {
		const proc = Bun.spawn(
			[Bun.argv[0], runnerPath, '--no-timings', '--correctness-boundary-json'],
			{
				cwd: process.cwd(),
				stderr: 'pipe',
				stdout: 'pipe',
			},
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])

		expect(exitCode, stderr).toBe(0)
		expect(stderr.trim()).toBe('')
		const packet = JSON.parse(stdout) as {
			readonly ownerLoop?: string
			readonly status?: string
			readonly releaseGate?: string
			readonly headlineClaimsAllowed?: boolean
			readonly ownerApprovalRequired?: boolean
			readonly artifact?: string
			readonly gateId?: string
			readonly allCurrentEvidencePresent?: boolean
			readonly missingFeatureNames?: readonly string[]
			readonly approvalChecklist?: readonly {
				readonly gateId?: string
				readonly ownerLoop?: string
			}[]
			readonly featureChecks?: readonly {
				readonly feature?: string
				readonly evidencePresent?: boolean
				readonly forbiddenWording?: string
			}[]
			readonly validationCommands?: readonly string[]
			readonly forbiddenShortcuts?: readonly string[]
		}
		expect(packet).toMatchObject({
			ownerLoop: 'correctness',
			status: 'owner-decision-required',
			releaseGate: 'blocked-by-publication-policy',
			headlineClaimsAllowed: false,
			ownerApprovalRequired: true,
			artifact: 'package-action-proof',
			gateId: 'unsupported-feature-boundary',
			allCurrentEvidencePresent: true,
			missingFeatureNames: [],
		})
		expect(packet.approvalChecklist?.map((item) => `${item.ownerLoop}/${item.gateId}`)).toEqual([
			'correctness/unsupported-feature-boundary',
		])
		expect(packet.featureChecks?.map((item) => item.feature)).toEqual([
			'digital-signatures',
			'calc-chain',
			'chart-drawing-sidecars',
			'macros-activex',
			'unknown-parts',
			'encrypted-files',
			'streaming-scope',
		])
		expect(packet.featureChecks?.every((item) => item.evidencePresent === true)).toBe(true)
		expect(packet.featureChecks?.[0].forbiddenWording).toContain('re-signs')
		expect(packet.validationCommands).toEqual([
			'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json',
			'bun run fixtures/benchmarks/package-action-proof.ts --no-timings --json',
		])
		expect(packet.forbiddenShortcuts?.join('\n')).toContain('semantic support')
		expect(stdout).not.toContain('"claimBlockerBoard"')
		expect(stdout).not.toContain('"fixturePolicy"')
		expect(stdout).not.toContain('"deferredClaims"')
		expect(stdout).not.toContain('"qssLeapfrogReleaseMatrix"')
	})

	test('emits a compact performance boundary packet JSON mode', async () => {
		const proc = Bun.spawn(
			[Bun.argv[0], runnerPath, '--no-timings', '--performance-boundary-json'],
			{
				cwd: process.cwd(),
				stderr: 'pipe',
				stdout: 'pipe',
			},
		)
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])

		expect(exitCode, stderr).toBe(0)
		expect(stderr.trim()).toBe('')
		const packet = JSON.parse(stdout) as {
			readonly ownerLoop?: string
			readonly status?: string
			readonly releaseGate?: string
			readonly headlineClaimsAllowed?: boolean
			readonly ownerApprovalRequired?: boolean
			readonly broadSpeedClaimAllowed?: boolean
			readonly benchmarkBlocker?: {
				readonly artifactId?: string
				readonly path?: string
				readonly validationCommand?: string
				readonly nextAction?: string
				readonly benchmarkCommands?: readonly string[]
				readonly acceptanceEvidence?: readonly string[]
				readonly stopCondition?: string
			}
			readonly approvalChecklist?: readonly {
				readonly gateId?: string
				readonly ownerLoop?: string
			}[]
			readonly validationCommands?: readonly string[]
			readonly forbiddenShortcuts?: readonly string[]
		}
		expect(packet).toMatchObject({
			ownerLoop: 'performance',
			status: 'owner-decision-required',
			releaseGate: 'blocked-by-publication-policy',
			headlineClaimsAllowed: false,
			ownerApprovalRequired: true,
			broadSpeedClaimAllowed: false,
			benchmarkBlocker: {
				artifactId: 'performance-claim-baseline-matrix',
				path: 'docs/PERFORMANCE_CLAIM_BASELINE_MATRIX.md',
				validationCommand: 'bun test fixtures/benchmarks/performance-claim-baseline-matrix.test.ts',
			},
		})
		expect(packet.benchmarkBlocker?.nextAction).toContain('ClosedXML head-to-head read run')
		expect(packet.benchmarkBlocker?.nextAction).toContain('fastxlsx runner')
		expect(packet.benchmarkBlocker?.benchmarkCommands?.join('\n')).toContain(
			'competitive-scoreboard.ts <suite.json> --json --metric medianMs --require-profile xlsx-read-sota',
		)
		expect(packet.benchmarkBlocker?.acceptanceEvidence?.join('\n')).toContain('ClosedXML')
		expect(packet.benchmarkBlocker?.acceptanceEvidence?.join('\n')).toContain('not counted as wins')
		expect(packet.approvalChecklist?.map((item) => `${item.ownerLoop}/${item.gateId}`)).toEqual([
			'performance/release-latency-run',
			'performance/streaming-matrix-boundary',
		])
		expect(packet.validationCommands?.join('\n')).toContain(
			'bun test fixtures/benchmarks/performance-claim-baseline-matrix.test.ts',
		)
		expect(packet.forbiddenShortcuts?.join('\n')).toContain('fastest XLSX reader')
		expect(stdout).not.toContain('"claimBlockerBoard"')
		expect(stdout).not.toContain('"fixturePolicy"')
		expect(stdout).not.toContain('"deferredClaims"')
		expect(stdout).not.toContain('"qssLeapfrogReleaseMatrix"')
	})

	test('emits a compact research hygiene decision packet JSON mode', async () => {
		const proc = Bun.spawn([Bun.argv[0], runnerPath, '--no-timings', '--research-hygiene-json'], {
			cwd: process.cwd(),
			stderr: 'pipe',
			stdout: 'pipe',
		})
		const [stdout, stderr, exitCode] = await Promise.all([
			new Response(proc.stdout).text(),
			new Response(proc.stderr).text(),
			proc.exited,
		])

		expect(exitCode, stderr).toBe(0)
		expect(stderr.trim()).toBe('')
		const packet = JSON.parse(stdout) as {
			readonly ownerLoops?: readonly string[]
			readonly status?: string
			readonly releaseGate?: string
			readonly headlineClaimsAllowed?: boolean
			readonly ownerApprovalRequired?: boolean
			readonly claim?: string
			readonly workBlockDisposition?: string
			readonly dirtyInventoryCommand?: string
			readonly inventorySnapshot?: {
				readonly command?: string
				readonly status?: string
				readonly decision?: string
				readonly dirtyPathCount?: number
				readonly statusCodeCounts?: Record<string, number>
				readonly rootCounts?: Record<string, number>
				readonly untrackedDirectoryCount?: number
				readonly modifiedFileCount?: number
				readonly unclassifiedEntries?: readonly {
					readonly statusCode?: string
					readonly path?: string
					readonly classification?: string
				}[]
				readonly boundary?: string
			}
			readonly validationCommands?: readonly string[]
			readonly ownerFiles?: readonly string[]
			readonly classificationBuckets?: readonly {
				readonly bucket?: string
				readonly requirement?: string
				readonly forbiddenShortcut?: string
			}[]
			readonly failureEvidence?: readonly string[]
			readonly acceptanceCriteria?: string
			readonly allowedWording?: string
			readonly forbiddenWording?: readonly string[]
			readonly qssContrast?: readonly string[]
			readonly nextOwnerAction?: string
			readonly stopCondition?: string
			readonly boundary?: string
		}
		expect(packet).toMatchObject({
			ownerLoops: ['product', 'release'],
			status: 'claim-downgrade-do-not-promote',
			releaseGate: 'blocked-by-publication-policy',
			headlineClaimsAllowed: false,
			ownerApprovalRequired: true,
			claim: 'research-surface-hygiene',
			workBlockDisposition: 'claim-downgrade-do-not-promote',
			dirtyInventoryCommand:
				'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
		})
		expect(packet.inventorySnapshot).toMatchObject({
			command: packet.dirtyInventoryCommand,
			status: 'inventory-collected',
			boundary: expect.stringContaining(
				'Inventory snapshot is current git-status routing evidence',
			),
		})
		expect(typeof packet.inventorySnapshot?.dirtyPathCount).toBe('number')
		expect(['owner-classification-required', 'no-unclassified-paths-currently-visible']).toContain(
			packet.inventorySnapshot?.decision,
		)
		expect(packet.inventorySnapshot?.dirtyPathCount).toBe(
			packet.inventorySnapshot?.unclassifiedEntries?.length ?? 0,
		)
		const statusTotal = Object.values(packet.inventorySnapshot?.statusCodeCounts ?? {}).reduce(
			(sum, count) => sum + count,
			0,
		)
		const rootTotal = Object.values(packet.inventorySnapshot?.rootCounts ?? {}).reduce(
			(sum, count) => sum + count,
			0,
		)
		expect(statusTotal).toBe(packet.inventorySnapshot?.dirtyPathCount)
		expect(rootTotal).toBe(packet.inventorySnapshot?.dirtyPathCount)
		expect(packet.inventorySnapshot?.untrackedDirectoryCount).toBe(
			(packet.inventorySnapshot?.unclassifiedEntries ?? []).filter(
				(entry) => entry.statusCode === '??' && entry.path?.endsWith('/'),
			).length,
		)
		expect(packet.inventorySnapshot?.modifiedFileCount).toBe(
			(packet.inventorySnapshot?.unclassifiedEntries ?? []).filter(
				(entry) => entry.statusCode === 'M',
			).length,
		)
		for (const entry of packet.inventorySnapshot?.unclassifiedEntries ?? []) {
			expect(entry.path).toMatch(
				/^(research|scripts\/ascend-loop-manager\.ts|tmp\/ascend-loop-manager)/,
			)
			expect(entry.statusCode?.length).toBeGreaterThan(0)
			expect(entry.classification).toBe('unclassified-owner-decision-required')
		}
		expect(packet.validationCommands).toEqual([
			'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
			'bun run fixtures/benchmarks/release-proof-index.ts --no-timings --research-hygiene-json',
			'bun test fixtures/benchmarks/release-proof-index.test.ts',
		])
		expect(packet.ownerFiles).toEqual(
			expect.arrayContaining([
				'research/',
				'research/experiments/',
				'scripts/ascend-loop-manager.ts',
				'tmp/ascend-loop-manager/',
			]),
		)
		expect(packet.classificationBuckets?.map((entry) => entry.bucket)).toEqual([
			'accepted-evidence',
			'active-owner-blocker',
			'archive-only',
		])
		expect(packet.failureEvidence?.join('\n')).toContain('Owner-classified inventory')
		expect(packet.acceptanceCriteria).toContain('classifies each untriaged path')
		expect(packet.allowedWording).toContain('Do not promote research-surface-hygiene')
		expect(packet.forbiddenWording?.join('\n')).toContain('Do not cite `research/` or `tmp/`')
		expect(packet.qssContrast?.join('\n')).toContain('QSS comparison is blocked')
		expect(packet.nextOwnerAction).toContain('git status --short research')
		expect(packet.nextOwnerAction).toContain('--research-hygiene-json')
		expect(packet.stopCondition).toContain(
			'accepted evidence, active owner blocker, or archive-only',
		)
		expect(packet.boundary).toContain('Compact research hygiene decision packet')
		expect(stdout).not.toContain('"claimBlockerBoard"')
		expect(stdout).not.toContain('"fixturePolicy"')
		expect(stdout).not.toContain('"qssLeapfrogReleaseMatrix"')
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
		expect(markdown).toContain('## Release Decision Board')
		expect(markdown).toContain('Do not promote yet:')
		expect(markdown).toContain('agent-view-budget-proof.test.ts')
		expect(markdown).toContain('agent-view-recovery-proof.test.ts')
		expect(markdown).toContain('Do not claim exact model-token counts')
		expect(markdown).toContain('viewport-patch-proof.test.ts')
		expect(markdown).toContain('Do not claim collaboration, sync, CRDT')
		expect(markdown).toContain('packages/sdk/src/agent-workflow.test.ts')
		expect(markdown).toContain('Do not claim autonomous correctness')
		expect(markdown).toContain('trace payload size')
		expect(markdown).toContain(
			'git status --short research scripts/ascend-loop-manager.ts tmp/ascend-loop-manager',
		)
		expect(markdown).toContain('Do not cite `research/` or `tmp/` files')
		expect(markdown).toContain(
			'| Rank | Claim | Evidence we have | Evidence missing | QSS contrast | Allowed wording | Forbidden wording | Next owner action | Owner decision artifact | Headline claim allowed | Implementation promotion allowed | Exact proof | Must not claim | A+ blocking owner action | Boundary |',
		)
		expect(markdown).toContain('| 1 | safe unknown workbook opening |')
		expect(markdown).toContain('performance/performance-claim-baseline-matrix')
		expect(markdown).toContain('correctness/excel-behavior-compatibility-matrix')
		expect(markdown).toContain('QSS likely does well:')
		expect(markdown).toContain('Ascend proven today:')
		expect(markdown).toContain('pre-hydration package-feature routing')
		expect(markdown).toContain('| 2 | auditable package-part mutation |')
		expect(markdown).toContain('per-part package action accounting')
		expect(markdown).toContain('performance/package-action-streaming-matrix-evidence')
		expect(markdown).toContain('release-latency-run: validation-run')
		expect(markdown).toContain('provenance-boundary: publication-policy')
		expect(markdown).toContain('Release Packageability Evidence')
		expect(markdown).toContain('bun run release:apps:smoke')
		expect(markdown).toContain('bun run release:rc:gate')
		expect(markdown).toContain('both compact proof commands')
		expect(markdown).toContain('offline verification policy')
		expect(markdown).toContain('local tarball install')
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
		expect(markdown).toContain(
			'Claim blocker board: safe-open-proof/performance=release-latency-run',
		)
		expect(markdown).toContain(
			'package-action-proof/release=provenance-boundary,compact-report-publication-policy',
		)
		expect(markdown).toContain('## Claim Blocker Board')
		expect(markdown).toContain(
			'| Artifact | Claim | Owner loop | Blockers | Next steps | Acceptance evidence | Forbidden shortcuts | Boundary |',
		)
		expect(markdown).toContain('| safe-open-proof | safe unknown workbook opening | product')
		expect(markdown).toContain('| package-action-proof | auditable package-part mutation | release')
		expect(markdown).toContain('## Next Owner Actions')
		expect(markdown).toContain('## Fixture Policy')
		expect(markdown).toContain('Current decision: owner-approval-required')
		expect(markdown).toContain(
			'safe-open-proof=`bun run fixtures/benchmarks/safe-open-fixture-scan.ts --json`',
		)
		expect(markdown).toContain(
			'package-action-proof=`bun run fixtures/benchmarks/package-action-fixture-scan.ts --json`',
		)
		expect(markdown).toContain('safe-open-proof=signed,malformed')
		expect(markdown).toContain('package-action-proof=signature-invalidation-drop')
		expect(markdown).toContain('edge case is package-topology-only')
		expect(markdown).toContain('Public binary fixtures are required when:')
		expect(markdown).toContain('Safe-open generated case acceptance checklist:')
		expect(markdown).toContain('| signed | generated-edge-package')
		expect(markdown).toContain('keeps-public-edge-fixtures-missing-until-owner-approval')
		expect(markdown).toContain('Package-action generated case acceptance checklist:')
		expect(markdown).toContain('| signature-invalidation-drop | generated-edge-package')
		expect(markdown).toContain('keeps-edge-fixture-policy-missing-until-owner-approval')
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
		expect(markdown).toContain('signatureOrUnknownMatches=1')
		expect(markdown).toContain('External fixture candidates:')
		expect(markdown).toContain('excelforge-book1-unknown-part')
		expect(markdown).toContain('satisfies-unknown-part-only')
		expect(markdown).toContain('excelforge-book1-unknown-part-mutation')
		expect(markdown).toContain('| package-action-proof | edge-fixture-policy')
		expect(markdown).toContain('missingReplacementFeatures=signaturePackage')
		expect(markdown).toContain('Fixture acquisition plan:')
		expect(markdown).toContain('signed-package')
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
		expect(markdown).toContain('Safe-open latency validation evidence:')
		expect(markdown).toContain('Status: timed-evidence-absent-owner-run-required')
		expect(markdown).toContain('Run profile: safe-open-release-latency-owner-review')
		expect(markdown).toContain(
			'Run profile command: `bun run fixtures/benchmarks/safe-open-proof.ts --repeat 10 --warmup 3 --json`',
		)
		expect(markdown).toContain('Run profile satisfied: false')
		expect(markdown).toContain('repeat 1 below profile minimum 10')
		expect(markdown).toContain('Run profile minimums: repeat 10, warmup 3')
		expect(markdown).toContain(
			'Run profile public cases: clean,formula-heavy,macro,pivot,activex,chart',
		)
		expect(markdown).toContain('Run profile CV guard: publicOpenPlanCv <= 0.25')
		expect(markdown).toContain('Timed case count: 0')
		expect(markdown).toContain('Timing environment captured: false')
		expect(markdown).toContain('Public open-plan p95 ms: {}')
		expect(markdown).toContain('Public open-plan CV: {}')
		expect(markdown).toContain('Public full-open p95 ms: {}')
		expect(markdown).toContain('Public full-open CV: {}')
		expect(markdown).toContain('Release claim allowed: false')
		expect(markdown).toContain('Threshold claim allowed: false')
		expect(markdown).toContain('Missing latency policy requirements:')
		expect(markdown).toContain('tracked-clean release environment')
		expect(markdown).toContain('non-threshold release wording')
		expect(markdown).toContain('## Streaming Matrix Evidence')
		expect(markdown).toContain('Status: representative-proof-present-owner-approval-required')
		expect(markdown).toContain('Covered action kinds: passthrough,regenerate,add,drop')
		expect(markdown).toContain('Missing action kinds: error')
		expect(markdown).toContain(
			'Covered cases: docprops-passthrough,add-sheet-part,calc-chain-drop,macro-passthrough,chart-sidecar-accounting',
		)
		expect(markdown).toContain('Non-streaming cases: regenerate-existing-sheet')
		expect(markdown).toContain('Public non-streaming cases: unknown-part-error')
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
		expect(markdown).toContain('Correctness/trust completeness boundary:')
		expect(markdown).toContain('Status: boundary-pinned-owner-scope')
		expect(markdown).toContain('Does not close gates: product,performance,release')
		expect(markdown).toContain('| Out-of-scope class | Promote only when | Owner action |')
		expect(markdown).toContain('| Broad formula function coverage |')
		expect(markdown).toContain('does not approve headline release claims')
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
			'| Rank | Artifact | Gate | Owner loop | Priority | Next step | Validation command | Acceptance evidence | Forbidden shortcut |',
		)
		expect(markdown).toContain(
			'| 10 | package-action-proof | edge-fixture-policy | product | claim-evidence | owner-decision-or-fixture-replacement | `bun run fixtures/benchmarks/package-action-fixture-scan.ts --json` | Product accepts disclosed generated',
		)
		expect(markdown).toContain(
			'| 50 | package-action-proof | provenance-boundary | release | publication-policy | publication-policy | `bun run fixtures/benchmarks/release-proof-index.ts --no-timings --owner-handoffs-json` | Release approves local package-action proof wording below SLSA',
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
		expect(markdown).toContain('generatedCases=signed,malformed')
		expect(markdown).toContain('deterministicGenerated=signed,malformed')
		expect(markdown).toContain('generatedDigests=signed:')
		expect(markdown).toContain('safe-open-proof.ts --no-timings --json')
		expect(markdown).toContain('safe-open-proof.ts --no-timings --compact-json')
		expect(markdown).toContain('package-action-proof.ts --no-timings --compact-json')
		expect(markdown).toContain('SLSA')
		expect(markdown).toContain('Attestation: false')
		expect(markdown).toContain('safe unknown workbook opening')
		expect(markdown).toContain('auditable package-part mutation')
		expect(markdown).toContain('streamingProofCases=5')
		expect(markdown).toContain('Excluded Evidence')
		expect(markdown).toContain('practical-latency-contracts')
		expect(markdown).toContain('tracked-clean run')
		expect(markdown).toContain('research surface as release evidence')
		expect(markdown).toContain('Untriaged research files are not release evidence')
		expect(markdown).toContain('Ranked Claim Portfolio')
		expect(markdown).toContain(
			'| Rank | Claim | Status | North Star link | Owner loops | Handoff decision | Proof command | Kill criterion | Boundary |',
		)
		expect(markdown).toContain('| 1 | safe unknown workbook opening | claim-wording-allowed-today')
		expect(markdown).toContain(
			'| 2 | auditable package-part mutation | claim-wording-allowed-today',
		)
		expect(markdown).toContain('| 6 | release proof bundle | needs-one-more-fold-in')
		expect(markdown).toContain('| 8 | property-style journal laws | speculative-do-not-promote')
		expect(markdown).toContain('top-implementation-handoff')
		expect(markdown).toContain('proof-packaging-only')
		expect(markdown).toContain('do-not-promote-yet')
		expect(markdown).toContain('Deferred Claims')
		expect(markdown).toContain('formula language-service primitives')
		expect(markdown).toContain('edit-producing rename is frozen')
		expect(markdown).toContain('formula-assist-proof.ts')
		expect(markdown).toContain('Do not claim edit-producing rename')
		expect(markdown).toContain('formula-corpus-correctness.test.ts')
		expect(markdown).toContain('Do not claim Excel-compatible formulas')
		expect(markdown).toContain('journal-law-proof.test.ts')
		expect(markdown).toContain('Do not claim property-based testing')
		expect(markdown).toContain('columnar-sidecar.test.ts')
		expect(markdown).toContain('Do not claim a production cache')
		expect(markdown).toContain('columnar scan sidecars')
		expect(markdown).toContain('do-not-promote-yet')
	})
})
