import { afterEach, describe, expect, test } from 'bun:test'
import { Buffer } from 'node:buffer'
import { existsSync, mkdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { inspectXlsxPackageGraph } from '@ascend/io-xlsx'
import { createZip, encode } from '../../io-xlsx/src/writer/zip.ts'
import { makeEmbeddedChartXlsx, makeXlsx } from '../../io-xlsx/test/helpers.ts'
import {
	AscendWorkbook,
	auditLossPolicy,
	auditPackageGraphIntegrity,
	commitAgentPlan,
	createAgentPlan,
} from './index.ts'

const TEMP_DIR = join(tmpdir(), `ascend-agent-workflow-${process.pid}`)

afterEach(() => {
	if (existsSync(TEMP_DIR)) rmSync(TEMP_DIR, { recursive: true, force: true })
})

describe('agent workflow loss audit', () => {
	test('docProps-only package preservation does not require lossy-write approval', () => {
		const cleanDocPropsAudit = auditLossPolicy([
			{
				feature: 'preservedDocumentProperties',
				tier: 'preserved',
				count: 3,
				locations: ['docProps/core.xml', 'docProps/app.xml', 'docProps/custom.xml'],
			},
		])
		expect(cleanDocPropsAudit.ok).toBe(true)
		expect(cleanDocPropsAudit.blockedFeatures).toHaveLength(0)

		const legacyDocPropsAudit = auditLossPolicy([
			{
				feature: 'preservedOther',
				tier: 'preserved',
				count: 3,
				locations: ['docProps/core.xml', 'docProps/app.xml', 'docProps/custom.xml'],
			},
		])
		expect(legacyDocPropsAudit.ok).toBe(true)
		expect(legacyDocPropsAudit.blockedFeatures).toHaveLength(0)

		const calcChainAudit = auditLossPolicy([
			{
				feature: 'preservedCalcChain',
				tier: 'preserved',
				count: 1,
				locations: ['xl/calcChain.xml'],
			},
		])
		expect(calcChainAudit.ok).toBe(true)
		expect(calcChainAudit.blockedFeatures).toHaveLength(0)

		const unknownPackageAudit = auditLossPolicy([
			{
				feature: 'preservedOther',
				tier: 'preserved',
				count: 1,
				locations: ['xl/custom/custom1.xml'],
			},
		])
		expect(unknownPackageAudit.ok).toBe(false)
		expect(unknownPackageAudit.blockedFeatures[0]?.feature).toBe('preservedOther')
		expect(unknownPackageAudit.blockedPackageParts).toHaveLength(0)

		const packageGraphAudit = auditLossPolicy(
			[
				{
					feature: 'preservedOther',
					tier: 'preserved',
					count: 1,
					locations: ['xl/custom/custom1.xml'],
				},
			],
			[],
			inspectXlsxPackageGraph(makePreservedCustomXlsx()),
		)
		expect(packageGraphAudit.blockedPackageParts).toEqual([
			expect.objectContaining({
				partPath: 'xl/custom/custom1.xml',
				featureFamily: 'preservedOther',
				preservationPolicy: 'unknown-review-required',
				ownerScope: 'unknown',
				bytePreservationExpected: false,
			}),
		])
		const fidelityAudit = auditPackageGraphIntegrity(
			inspectXlsxPackageGraph(makePreservedCustomXlsx()),
		)
		expect(fidelityAudit.ok).toBe(false)
		expect(fidelityAudit.issues).toContainEqual(
			expect.objectContaining({
				code: 'package_feature_classification',
				partPath: 'xl/custom/custom1.xml',
			}),
		)
	})

	test('plans report blocked preserved features and commits require explicit allow-loss', async () => {
		const input = join(TEMP_DIR, 'preserved.xlsx')
		const output = join(TEMP_DIR, 'out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makePreservedCustomXlsx())
		const ops = [{ op: 'setCells' as const, sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }]

		const plan = await createAgentPlan(input, ops)
		expect(plan.lossAudit.ok).toBe(false)
		expect(plan.lossAudit.blockedFeatures[0]?.feature).toBe('preservedOther')
		expect(plan.packageGraphAudit.ok).toBe(false)
		expect(plan.packageGraphAudit.issues).toContainEqual(
			expect.objectContaining({
				code: 'package_feature_classification',
				partPath: 'xl/custom/custom1.xml',
			}),
		)
		expect(plan.lossAudit.blockedPackageParts).toEqual([
			expect.objectContaining({
				partPath: 'xl/custom/custom1.xml',
				featureFamily: 'preservedOther',
				preservationPolicy: 'unknown-review-required',
				reason: expect.stringContaining('cannot classify'),
			}),
		])
		expect(plan.trace.kind).toBe('plan')
		expect(plan.trace.traceDigest).toMatch(/^[a-f0-9]{64}$/)
		const lossPhase = plan.trace.phases.find((phase) => phase.phase === 'loss-audit')
		expect(lossPhase?.status).toBe('blocked')
		const packageGraphPhase = plan.trace.phases.find(
			(phase) => phase.phase === 'package-graph-audit',
		)
		expect(packageGraphPhase?.status).toBe('warning')
		expect(lossPhase?.details).toMatchObject({
			blockedPackageParts: [
				expect.objectContaining({
					partPath: 'xl/custom/custom1.xml',
					preservationPolicy: 'unknown-review-required',
				}),
			],
		})
		expect(plan.needsApproval).toBe(true)
		expect(plan.approvals[0]?.kind).toBe('lossy-write')
		expect(plan.modelOutput.blocked).toBe(true)
		expect(plan.modelOutput.counts.packageGraphIssues).toBe(1)
		expect(plan.modelOutput.nextActions.join('\n')).toContain('approval')

		await expect(commitAgentPlan(input, ops, { output })).rejects.toThrow(
			'Commit requires explicit approval',
		)

		const committed = await commitAgentPlan(input, ops, {
			output,
			allowLoss: ['preservedOther'],
		})
		expect(committed.lossAudit.ok).toBe(true)
		expect(committed.packageGraphAudit.ok).toBe(false)
		expect(committed.postWrite.packageGraphAudit.ok).toBe(false)
		expect(committed.postWrite.packageGraphAudit.issues).toContainEqual(
			expect.objectContaining({
				code: 'package_feature_classification',
				partPath: 'xl/custom/custom1.xml',
			}),
		)
		expect(committed.outputSha256).toMatch(/^[a-f0-9]{64}$/)
		expect(committed.postWrite.valid).toBe(true)
		expect(committed.postWrite.outputSha256).toBe(committed.outputSha256)
		expect(committed.trace.kind).toBe('commit')
		expect(committed.trace.outputSha256).toBe(committed.outputSha256)
		expect(committed.trace.phases.find((phase) => phase.phase === 'post-write')?.status).toBe(
			'warning',
		)
		expect(committed.modelOutput.counts.postWritePackageGraphIssues).toBeGreaterThan(0)
		expect(committed.modelOutput.nextActions.join('\n')).toContain(
			'postWrite.packageGraphAudit.issues',
		)
		expect(committed.modelOutput.blocked).toBe(false)
		expect(committed.modelOutput.digests.traceDigest).toBe(committed.trace.traceDigest)
	})

	test('plans expose digital signature invalidation package policy', async () => {
		const input = join(TEMP_DIR, 'signed.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeSignedXlsx())

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
		])
		expect(plan.lossAudit.ok).toBe(false)
		expect(plan.lossAudit.blockedFeatures).toContainEqual(
			expect.objectContaining({ feature: 'preservedSignature' }),
		)
		expect(plan.lossAudit.blockedPackageParts).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					partPath: '_xmlsignatures/origin.sigs',
					featureFamily: 'preservedSignature',
					preservationPolicy: 'invalidate-on-edit',
					sourceRelationshipPart: '_rels/.rels',
					sourceRelationshipId: 'rIdSignatureOrigin',
					reason: expect.stringContaining('invalidate'),
				}),
				expect.objectContaining({
					partPath: '_xmlsignatures/sig1.xml',
					featureFamily: 'preservedSignature',
					preservationPolicy: 'invalidate-on-edit',
					sourceRelationshipPart: '_xmlsignatures/_rels/origin.sigs.rels',
					sourceRelationshipId: 'rIdSignature',
				}),
			]),
		)
		expect(plan.preservation.skippedCapsules).toEqual(
			expect.arrayContaining(['_xmlsignatures/origin.sigs', '_xmlsignatures/sig1.xml']),
		)
		expect(plan.writePolicy.summary.invalidatedSignatures).toBe(2)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'signature-invalidation',
				severity: 'warning',
				partPaths: expect.arrayContaining([
					'_xmlsignatures/origin.sigs',
					'_xmlsignatures/sig1.xml',
				]),
			}),
		)
		expect(plan.trace.phases.find((phase) => phase.phase === 'write-policy')?.status).toBe(
			'warning',
		)
	})

	test('clean workbooks commit without allow-loss', async () => {
		const input = join(TEMP_DIR, 'clean.xlsx')
		const output = join(TEMP_DIR, 'clean-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		await wb.save(input)

		const committed = await commitAgentPlan(
			input,
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] }],
			{ output },
		)
		expect(committed.lossAudit.ok).toBe(true)
		expect(committed.writePolicy.ok).toBe(true)
		expect(committed.writePolicy.diagnostics.every((entry) => entry.severity === 'info')).toBe(true)
		expect(
			committed.writePolicy.diagnostics.some((entry) => entry.code.startsWith('external-link')),
		).toBe(false)
		expect(
			committed.writePolicy.diagnostics.some(
				(entry) => entry.code === 'analytics-pivot-refresh-risk',
			),
		).toBe(false)
		expect(committed.writePolicy.diagnostics.some((entry) => entry.code.includes('comment'))).toBe(
			false,
		)
		expect(committed.packageGraphAudit.ok).toBe(true)
		expect(committed.postWrite.valid).toBe(true)
		expect(committed.postWrite.packageGraphAudit.ok).toBe(true)
		expect(committed.postWrite.packageGraphAudit.policy).toBe('safe-edit-roundtrip')
		expect(committed.modelOutput.counts.postWritePackageGraphIssues).toBe(0)
		expect(committed.postWrite.reopened).toBe(true)
		expect(committed.postWrite.check.valid).toBe(true)
		expect(committed.trace.artifacts.map((artifact) => artifact.name)).toContain('apply')
		expect(committed.trace.artifacts.map((artifact) => artifact.name)).toContain('postWrite')
		expect(committed.modelOutput.counts.operations).toBe(1)
	})

	test('plans explain calc chain preservation versus formula-topology invalidation', async () => {
		const input = join(TEMP_DIR, 'calc-chain.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeCalcChainXlsx())

		const valueEdit = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 2 }] },
		])
		expect(valueEdit.writePolicy.summary.calcChainPolicy).toBe('preserved')
		expect(valueEdit.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'calc-chain-preserved',
				severity: 'info',
				partPaths: ['xl/calcChain.xml'],
			}),
		)

		const formulaEdit = await createAgentPlan(input, [
			{ op: 'setFormula', sheet: 'Sheet1', ref: 'B1', formula: '=A1*3' },
		])
		expect(formulaEdit.preservation.skippedCapsules).toContain('xl/calcChain.xml')
		expect(formulaEdit.writePolicy.summary.calcChainPolicy).toBe('discarded-for-formula-topology')
		expect(formulaEdit.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'calc-chain-discarded',
				severity: 'warning',
				preservationPolicy: 'discard-on-recalc',
			}),
		)
	})

	test('plans explain external link package binding risk', async () => {
		const input = join(TEMP_DIR, 'external-link-missing-binding.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeExternalLinkMissingBindingXlsx())

		const plan = await createAgentPlan(input, [
			{
				op: 'setFormula',
				sheet: 'Sheet1',
				ref: 'A1',
				formula: '=SUM([1]FY26!B2:B10)',
			},
			{
				op: 'setDefinedName',
				name: 'BudgetSource',
				ref: '[1]FY26!A1:D10',
			},
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'C2:C20',
				rule: { type: 'list', formula1: '[1]FY26!$A$1:$A$4' },
			},
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'D2:D20',
				rule: {
					type: 'expression',
					formula: '=SUM([1]FY26!D2:D20)>0',
					dataBar: { cfvo: [{ type: 'formula', value: '[1]FY26!$D$1' }] },
				},
			},
		])

		expect(plan.writePolicy.summary.externalReferences).toBe(1)
		expect(plan.writePolicy.summary.externalReferenceBindingIssues).toBe(1)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'external-link-dependency',
				severity: 'warning',
				partPaths: ['xl/externalLinks/externalLink1.xml'],
				packageParts: [
					expect.objectContaining({
						partPath: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
						preservationPolicy: 'preserve-exact',
					}),
				],
				details: expect.objectContaining({
					operationScoped: true,
					relatedOperations: expect.arrayContaining([
						expect.objectContaining({
							operationIndex: 0,
							op: 'setFormula',
							sourceKind: 'cellFormula',
							sourceRef: 'Sheet1!A1',
							formula: '=SUM([1]FY26!B2:B10)',
							references: ['[1]FY26!B2:B10'],
						}),
						expect.objectContaining({
							operationIndex: 1,
							op: 'setDefinedName',
							sourceKind: 'definedName',
							name: 'BudgetSource',
							formula: '[1]FY26!A1:D10',
							references: ['[1]FY26!A1:D10'],
						}),
						expect.objectContaining({
							operationIndex: 2,
							op: 'setDataValidation',
							sourceKind: 'dataValidation',
							sheetName: 'Sheet1',
							sourceRef: 'Sheet1!C2:C20',
							range: 'C2:C20',
							formula: '[1]FY26!$A$1:$A$4',
							references: ['[1]FY26!$A$1:$A$4'],
						}),
						expect.objectContaining({
							operationIndex: 3,
							op: 'setConditionalFormat',
							sourceKind: 'conditionalFormat',
							sheetName: 'Sheet1',
							sourceRef: 'Sheet1!D2:D20',
							range: 'D2:D20',
							formula: '=SUM([1]FY26!D2:D20)>0',
							references: ['[1]FY26!D2:D20'],
						}),
						expect.objectContaining({
							operationIndex: 3,
							op: 'setConditionalFormat',
							sourceKind: 'conditionalFormat',
							sheetName: 'Sheet1',
							sourceRef: 'Sheet1!D2:D20',
							range: 'D2:D20',
							formula: '[1]FY26!$D$1',
							references: ['[1]FY26!$D$1'],
						}),
					]),
				}),
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'external-link-binding-risk',
				severity: 'warning',
				partPaths: ['xl/externalLinks/externalLink1.xml'],
				message: expect.stringContaining('0 fallback, 1 missing'),
				suggestedAction: expect.stringContaining('rewriteExternalLink'),
				details: expect.objectContaining({
					operationScoped: true,
					bindingIssueCounts: { fallback: 0, missing: 1 },
					relatedOperations: expect.arrayContaining([
						expect.objectContaining({ op: 'setDataValidation', sourceKind: 'dataValidation' }),
						expect.objectContaining({
							op: 'setConditionalFormat',
							sourceKind: 'conditionalFormat',
						}),
					]),
					externalLinks: [
						expect.objectContaining({
							bindingRisk: expect.objectContaining({
								status: 'missingPathRelationship',
								externalBookRelId: 'rIdMissing',
							}),
						}),
					],
					rewriteExternalLinkRecommendations: [
						expect.objectContaining({
							op: 'rewriteExternalLink',
							partPath: 'xl/externalLinks/externalLink1.xml',
							newTarget: '<new-target>',
						}),
					],
				}),
			}),
		)
		expect(plan.writePolicy.ok).toBe(false)
		expect(plan.trace.phases.find((phase) => phase.phase === 'write-policy')?.status).toBe(
			'warning',
		)
	})

	test('keeps unresolved external-link package graph failures visible to agents', async () => {
		const input = join(TEMP_DIR, 'external-link-missing-target.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeMissingExternalLinkTargetXlsx())

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] },
		])

		expect(plan.packageGraphAudit.issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_target',
				relationshipPartPath: 'xl/_rels/workbook.xml.rels',
				relationshipId: 'rIdExternal',
				featureFamily: 'preservedExternalLink',
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'external-link-package-risk',
				severity: 'warning',
				partPaths: expect.arrayContaining(['xl/externalLinks/missing.xml']),
				featureFamily: 'preservedExternalLink',
				details: expect.objectContaining({
					packageGraphIssues: expect.arrayContaining([
						expect.objectContaining({
							code: 'package_relationship_target',
							relationshipId: 'rIdExternal',
							featureFamily: 'preservedExternalLink',
						}),
					]),
				}),
			}),
		)
		expect(
			plan.writePolicy.diagnostics.some(
				(diagnostic) =>
					diagnostic.code === 'package-graph-audit-issue' &&
					diagnostic.partPaths?.includes('xl/externalLinks/missing.xml'),
			),
		).toBe(false)
	})

	test('plans report preserved active content without implying execution support', async () => {
		const input = join(TEMP_DIR, 'macro.xlsm')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeMacroXlsx())

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
		])
		expect(plan.lossAudit.blockedFeatures).toContainEqual(
			expect.objectContaining({ feature: 'preservedMacro' }),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'active-content-preserved',
				severity: 'warning',
				featureFamily: 'preservedMacro',
				partPaths: ['xl/vbaProject.bin'],
			}),
		)
	})

	test('dirty cell edits preserve visual sidecars without visual-edit noise', async () => {
		const input = join(TEMP_DIR, 'visual.xlsx')
		const output = join(TEMP_DIR, 'visual-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeVisualXlsx())

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] },
		])
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'visual-sidecar-preservation-risk',
				severity: 'info',
				partPaths: expect.arrayContaining([
					'xl/drawings/drawing1.xml',
					'xl/media/image1.png',
					'xl/charts/style1.xml',
					'xl/charts/colors1.xml',
				]),
				packageParts: expect.arrayContaining([
					expect.objectContaining({
						partPath: 'xl/drawings/drawing1.xml',
						featureFamily: 'preservedDrawing',
						ownerScope: 'drawing',
						preservationPolicy: 'preserve-exact',
						bytePreservationExpected: true,
						contentTypeSource: 'override',
						sourceRelationshipPart: 'xl/worksheets/_rels/sheet1.xml.rels',
						sourceRelationshipId: 'rIdDrawing1',
					}),
					expect.objectContaining({
						partPath: 'xl/media/image1.png',
						featureFamily: 'preservedMedia',
						ownerScope: 'drawing',
						contentTypeSource: 'default',
						sourceRelationshipPart: 'xl/drawings/_rels/drawing1.xml.rels',
						sourceRelationshipId: 'rIdImage1',
					}),
					expect.objectContaining({
						partPath: 'xl/charts/style1.xml',
						featureFamily: 'preservedChartStyle',
						ownerScope: 'unknown',
						contentTypeSource: 'override',
					}),
					expect.objectContaining({
						partPath: 'xl/charts/colors1.xml',
						featureFamily: 'preservedChartColor',
						ownerScope: 'unknown',
						contentTypeSource: 'override',
					}),
				]),
				details: expect.objectContaining({
					operationScoped: false,
					relatedOperations: [],
					chartSourceRefDrift: [],
					drawingmlVmlDrift: [],
				}),
			}),
		)
		expect(
			plan.writePolicy.diagnostics.some(
				(diagnostic) => diagnostic.code === 'visual-edit-preservation-risk',
			),
		).toBe(false)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'copied-through-package-parts',
				packageParts: expect.arrayContaining([
					expect.objectContaining({
						partPath: 'xl/charts/style1.xml',
						featureFamily: 'preservedChartStyle',
						preservationPolicy: 'preserve-exact',
					}),
					expect.objectContaining({
						partPath: 'xl/charts/colors1.xml',
						featureFamily: 'preservedChartColor',
						preservationPolicy: 'preserve-exact',
					}),
				]),
			}),
		)

		const committed = await commitAgentPlan(
			input,
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] }],
			{
				output,
				allowLoss: [
					'preservedDrawing',
					'preservedMedia',
					'preservedChartStyle',
					'preservedChartColor',
				],
			},
		)
		expect(committed.postWrite.packageGraphAudit.ok).toBe(true)
		const outputGraph = inspectXlsxPackageGraph(await Bun.file(output).bytes())
		expect(outputGraph.parts.map((part) => part.path)).toEqual(
			expect.arrayContaining([
				'xl/drawings/drawing1.xml',
				'xl/media/image1.png',
				'xl/charts/style1.xml',
				'xl/charts/colors1.xml',
			]),
		)
	})

	test('plans analytics pivot refresh risk for pivot slicer and timeline edits', async () => {
		const input = join(TEMP_DIR, 'analytics-refresh.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeAnalyticsRefreshXlsx())

		const plan = await createAgentPlan(input, [
			{
				op: 'setPivotFieldItem',
				pivotTable: 'PivotTable1',
				fieldIndex: 0,
				itemIndex: 1,
				hidden: true,
			},
			{
				op: 'setSlicerCacheItem',
				slicerCache: 'Slicer_Region',
				item: 0,
				selected: false,
			},
			{
				op: 'setTimelineRange',
				timelineCache: 'Timeline_Order_Date',
				startDate: '2024-01-01T00:00:00',
				endDate: '2024-03-31T00:00:00',
			},
		])

		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'visual-sidecar-preservation-risk',
				severity: 'warning',
				details: expect.objectContaining({
					operationScoped: true,
					analyticsPivotRefreshRisk: expect.arrayContaining([
						expect.objectContaining({ op: 'setPivotFieldItem' }),
						expect.objectContaining({ op: 'setSlicerCacheItem' }),
						expect.objectContaining({ op: 'setTimelineRange' }),
					]),
				}),
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'analytics-pivot-refresh-risk',
				severity: 'warning',
				partPaths: expect.arrayContaining([
					'xl/pivotTables/pivotTable1.xml',
					'xl/pivotCache/pivotCacheDefinition1.xml',
					'xl/pivotCache/pivotCacheRecords1.xml',
					'xl/slicerCaches/slicerCache1.xml',
					'xl/slicers/slicer1.xml',
					'xl/timelineCaches/timelineCache1.xml',
					'xl/timelines/timeline1.xml',
				]),
				packageParts: expect.arrayContaining([
					expect.objectContaining({
						partPath: 'xl/pivotTables/pivotTable1.xml',
						featureFamily: 'preservedPivot',
					}),
					expect.objectContaining({
						partPath: 'xl/slicerCaches/slicerCache1.xml',
						featureFamily: 'preservedSlicer',
					}),
					expect.objectContaining({
						partPath: 'xl/timelineCaches/timelineCache1.xml',
						featureFamily: 'preservedTimeline',
					}),
				]),
				suggestedAction: expect.stringContaining('Excel'),
				details: expect.objectContaining({
					operationScoped: true,
					analyticsPivotRefreshRisk: [
						expect.objectContaining({
							op: 'setPivotFieldItem',
							targetKind: 'pivotFieldItem',
							matchCount: 1,
							linkedPivotTableNames: ['PivotTable1'],
							partPaths: expect.arrayContaining([
								'xl/pivotTables/pivotTable1.xml',
								'xl/pivotCache/pivotCacheDefinition1.xml',
								'xl/pivotCache/pivotCacheRecords1.xml',
							]),
							selector: {
								pivotTable: 'PivotTable1',
								fieldIndex: 0,
								itemIndex: 1,
							},
							matches: [
								expect.objectContaining({
									pivotTable: 'PivotTable1',
									sheetName: 'PivotSheet',
									fieldExists: true,
									itemExists: true,
								}),
							],
						}),
						expect.objectContaining({
							op: 'setSlicerCacheItem',
							targetKind: 'slicerCacheItem',
							matchCount: 1,
							linkedPivotTableNames: ['PivotTable1'],
							partPaths: expect.arrayContaining([
								'xl/slicerCaches/slicerCache1.xml',
								'xl/slicers/slicer1.xml',
								'xl/pivotTables/pivotTable1.xml',
							]),
							selector: {
								slicerCache: 'Slicer_Region',
								item: 0,
							},
							matches: [
								expect.objectContaining({
									slicerCache: 'Slicer_Region',
									itemExists: true,
									linkedPivotTableNames: ['PivotTable1'],
									slicerPartPaths: ['xl/slicers/slicer1.xml'],
								}),
							],
						}),
						expect.objectContaining({
							op: 'setTimelineRange',
							targetKind: 'timelineRange',
							matchCount: 1,
							linkedPivotTableNames: ['PivotTable1'],
							partPaths: expect.arrayContaining([
								'xl/timelineCaches/timelineCache1.xml',
								'xl/timelines/timeline1.xml',
								'xl/pivotTables/pivotTable1.xml',
							]),
							selector: {
								timelineCache: 'Timeline_Order_Date',
								startDate: '2024-01-01T00:00:00',
								endDate: '2024-03-31T00:00:00',
							},
							matches: [
								expect.objectContaining({
									timelineCache: 'Timeline_Order_Date',
									linkedPivotTableNames: ['PivotTable1'],
									timelinePartPaths: ['xl/timelines/timeline1.xml'],
								}),
							],
						}),
					],
				}),
			}),
		)
	})

	test('plans image replacement with selector and sidecar recommendations', async () => {
		const input = join(TEMP_DIR, 'image-replace.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeVisualXlsx())

		const plan = await createAgentPlan(input, [
			{
				op: 'replaceImage',
				sheet: 'Sheet1',
				imageIndex: 0,
				contentBase64: Buffer.from('new-png-bytes').toString('base64'),
				contentType: 'image/png',
			},
		])

		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'visual-sidecar-preservation-risk',
				severity: 'warning',
				details: expect.objectContaining({ operationScoped: true }),
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'visual-edit-preservation-risk',
				severity: 'info',
				partPaths: expect.arrayContaining(['xl/drawings/drawing1.xml', 'xl/media/image1.png']),
				details: expect.objectContaining({
					relatedOperations: [
						expect.objectContaining({
							op: 'replaceImage',
							targetKind: 'image',
							selector: { imageIndex: 0 },
							matchCount: 1,
							matches: [
								expect.objectContaining({
									drawingPartPath: 'xl/drawings/drawing1.xml',
									targetPath: 'xl/media/image1.png',
									relId: 'rIdImage1',
								}),
							],
							recommendation: expect.stringContaining('unique image selector'),
						}),
					],
				}),
			}),
		)
	})

	test('plans chart source ref drift and safe setChartSeriesSource recommendation', async () => {
		const input = join(TEMP_DIR, 'chart-source-drift.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeEmbeddedChartXlsx({ sheetName: 'Data' }))

		const plan = await createAgentPlan(input, [
			{ op: 'insertRows', sheet: 'Data', at: 2, count: 1 },
		])

		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'chart-source-ref-drift-risk',
				severity: 'warning',
				partPaths: expect.arrayContaining([
					'xl/charts/chart1.xml',
					'xl/charts/style1.xml',
					'xl/charts/colors1.xml',
				]),
				suggestedAction: expect.stringContaining('setChartSeriesSource'),
				details: expect.objectContaining({
					chartSourceRefDrift: [
						expect.objectContaining({
							chartPartPath: 'xl/charts/chart1.xml',
							sheetName: 'Data',
							chartIndex: 0,
							seriesIndex: 0,
							sourceRefs: expect.arrayContaining([
								expect.objectContaining({
									sourceKind: 'categoryRef',
									ref: 'Data!$A$2:$A$4',
									referencedSheets: ['Data'],
								}),
								expect.objectContaining({
									sourceKind: 'valueRef',
									ref: 'Data!$B$2:$B$4',
									referencedSheets: ['Data'],
								}),
							]),
							relatedOperations: [
								expect.objectContaining({
									op: 'insertRows',
									sheetName: 'Data',
									rangeImpact: 'rows',
									at: 2,
									count: 1,
								}),
							],
							recommendation: expect.objectContaining({
								op: 'setChartSeriesSource',
								partPath: 'xl/charts/chart1.xml',
								chartIndex: 0,
								seriesIndex: 0,
							}),
						}),
					],
				}),
			}),
		)
	})

	test('plans DrawingML versus VML drift for visual edits on mixed sheets', async () => {
		const input = join(TEMP_DIR, 'mixed-drawings.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeMixedDrawingXlsx())

		const plan = await createAgentPlan(input, [
			{
				op: 'setDrawingText',
				sheet: 'Sheet1',
				drawingPartPath: 'xl/drawings/drawing1.xml',
				id: 2,
				text: 'Updated',
			},
		])

		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'drawingml-vml-drift-risk',
				severity: 'warning',
				partPaths: expect.arrayContaining([
					'xl/drawings/drawing1.xml',
					'xl/drawings/vmlDrawing1.vml',
				]),
				suggestedAction: expect.stringContaining('DrawingML plus VML'),
				details: expect.objectContaining({
					drift: [
						expect.objectContaining({
							op: 'setDrawingText',
							sheetName: 'Sheet1',
							drawingMlObjectCount: 1,
							vmlObjectCount: 1,
							recommendation: expect.stringContaining('separate package graphs'),
						}),
					],
				}),
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'visual-edit-preservation-risk',
				details: expect.objectContaining({
					relatedOperations: [
						expect.objectContaining({
							op: 'setDrawingText',
							targetKind: 'drawingText',
							matchCount: 1,
							matches: [
								expect.objectContaining({
									source: 'drawingml',
									editableText: true,
								}),
							],
						}),
					],
				}),
			}),
		)
	})

	test('plans group legacy comment XML and VML write-risk locations', async () => {
		const input = join(TEMP_DIR, 'legacy-comments.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().sheets[0]?.comments.set('B2', {
			text: 'Review this',
			author: 'Ada',
		})
		await wb.save(input)

		const plan = await createAgentPlan(input, [
			{ op: 'setComment', sheet: 'Sheet1', ref: 'B2', text: 'Review this again', author: 'Ada' },
		])

		expect(plan.writePolicy.summary.legacyCommentLocations).toBe(1)
		expect(plan.writePolicy.summary.commentIntegrityIssues).toBe(0)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'legacy-comment-preservation-risk',
				severity: 'warning',
				locations: ['Sheet1!B2'],
				partPaths: expect.arrayContaining(['xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml']),
				packageParts: expect.arrayContaining([
					expect.objectContaining({
						partPath: 'xl/comments1.xml',
						featureFamily: 'preservedComments',
					}),
					expect.objectContaining({
						partPath: 'xl/drawings/vmlDrawing1.vml',
						featureFamily: 'preservedVml',
					}),
				]),
				suggestedAction: expect.stringContaining('setComment'),
				details: expect.objectContaining({
					safeTextEdit: 'setComment',
					comments: [
						expect.objectContaining({
							sheetName: 'Sheet1',
							ref: 'B2',
							location: 'Sheet1!B2',
							author: 'Ada',
							hasLegacyDrawing: true,
						}),
					],
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'setComment',
							sheetName: 'Sheet1',
							ref: 'B2',
						}),
					],
					verifyIssues: [],
				}),
			}),
		)
	})

	test('plans flag structural edits that can shift legacy comment anchors', async () => {
		const input = join(TEMP_DIR, 'legacy-comment-structural-edit.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().sheets[0]?.comments.set('B2', {
			text: 'Review this',
			author: 'Ada',
		})
		await wb.save(input)

		const plan = await createAgentPlan(input, [
			{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 },
		])

		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'legacy-comment-preservation-risk',
				severity: 'warning',
				suggestedAction: expect.stringContaining('inspectSheet(sheet).comments'),
				details: expect.objectContaining({
					structuralEditVerification: expect.stringContaining('legacyDrawing/VML metadata'),
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'insertRows',
							sheetName: 'Sheet1',
							axis: 'row',
							range: '2:2',
							affectedLocations: ['Sheet1!B2'],
						}),
					],
				}),
			}),
		)
	})

	test('plans group threaded comment identity and persons sidecar write-risk locations', async () => {
		const input = join(TEMP_DIR, 'threaded-comments.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeThreadedCommentXlsx())

		const plan = await createAgentPlan(input, [
			{
				op: 'setThreadedComment',
				sheet: 'Sheet1',
				partPath: 'xl/threadedComments/threadedComment1.xml',
				threadedCommentId: 'tc2',
				text: 'Reviewed again',
			},
		])

		expect(plan.writePolicy.summary.threadedCommentLocations).toBe(2)
		expect(plan.writePolicy.summary.commentIntegrityIssues).toBe(0)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'threaded-comment-preservation-risk',
				severity: 'warning',
				locations: ['Sheet1!A1', 'Sheet1!A1'],
				partPaths: expect.arrayContaining([
					'xl/threadedComments/threadedComment1.xml',
					'xl/persons/person.xml',
				]),
				packageParts: expect.arrayContaining([
					expect.objectContaining({
						partPath: 'xl/threadedComments/threadedComment1.xml',
						featureFamily: 'preservedThreadedComments',
					}),
					expect.objectContaining({
						partPath: 'xl/persons/person.xml',
						featureFamily: 'preservedThreadedComments',
					}),
				]),
				suggestedAction: expect.stringContaining('setThreadedComment'),
				details: expect.objectContaining({
					safeTextEdit: 'setThreadedComment',
					threadedComments: [
						expect.objectContaining({
							sheetName: 'Sheet1',
							ref: 'A1',
							id: 'tc1',
							personId: '0',
							author: 'Ada Lovelace',
						}),
						expect.objectContaining({
							sheetName: 'Sheet1',
							ref: 'A1',
							id: 'tc2',
							parentId: 'tc1',
							personId: '1',
							author: 'Grace Hopper',
						}),
					],
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'setThreadedComment',
							sheetName: 'Sheet1',
							partPath: 'xl/threadedComments/threadedComment1.xml',
							threadedCommentId: 'tc2',
						}),
					],
					verifyIssues: [],
				}),
			}),
		)
	})

	test('plans flag structural edits that can shift threaded comment anchors', async () => {
		const input = join(TEMP_DIR, 'threaded-comment-structural-edit.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeThreadedCommentXlsx())

		const plan = await createAgentPlan(input, [
			{ op: 'deleteCols', sheet: 'Sheet1', at: 0, count: 1 },
		])

		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'threaded-comment-preservation-risk',
				severity: 'warning',
				suggestedAction: expect.stringContaining('inspectSheet(sheet).threadedComments'),
				details: expect.objectContaining({
					structuralEditVerification: expect.stringContaining('person metadata'),
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'deleteCols',
							sheetName: 'Sheet1',
							axis: 'column',
							range: 'A:A',
							affectedLocations: ['Sheet1!A1'],
						}),
					],
				}),
			}),
		)
	})

	test('plans include threaded comment verify issues in write-risk diagnostics', async () => {
		const input = join(TEMP_DIR, 'threaded-comments-missing-person.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeThreadedCommentXlsx({ includePersons: false }))

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 'ok' }] },
		])

		expect(plan.check.issues).toContainEqual(
			expect.objectContaining({
				rule: 'threaded-comment-integrity',
				refs: ['Sheet1!A1'],
			}),
		)
		expect(plan.writePolicy.summary.commentIntegrityIssues).toBe(3)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'threaded-comment-preservation-risk',
				severity: 'warning',
				partPaths: ['xl/threadedComments/threadedComment1.xml'],
				details: expect.objectContaining({
					verifyIssues: expect.arrayContaining([
						expect.objectContaining({
							rule: 'threaded-comment-integrity',
							refs: ['Sheet1!A1'],
							suggestedFix: expect.stringContaining('persons part'),
						}),
					]),
				}),
			}),
		)
	})

	test('plans attach legacy comment package graph relationship failures to comment risk', async () => {
		const input = join(TEMP_DIR, 'legacy-comments-missing-target.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeMissingLegacyCommentsTargetXlsx())

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] },
		])

		expect(plan.packageGraphAudit.issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_target',
				relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
				relationshipId: 'rIdComments',
				featureFamily: 'preservedComments',
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'legacy-comment-preservation-risk',
				details: expect.objectContaining({
					packageGraphIssues: expect.arrayContaining([
						expect.objectContaining({
							code: 'package_relationship_target',
							relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
							relationshipId: 'rIdComments',
						}),
					]),
					verifyIssues: expect.arrayContaining([
						expect.objectContaining({
							rule: 'legacy-comment-drawing-integrity',
							details: expect.objectContaining({
								kind: 'legacy-comments-relationship-missing-target',
							}),
						}),
					]),
				}),
			}),
		)
	})

	test('plans attach threaded comment package graph relationship failures to threaded risk', async () => {
		const input = join(TEMP_DIR, 'threaded-comments-missing-target.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeMissingThreadedCommentsTargetXlsx())

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] },
		])

		expect(plan.packageGraphAudit.issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_target',
				relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
				relationshipId: 'rIdThreaded',
				featureFamily: 'preservedThreadedComments',
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'threaded-comment-preservation-risk',
				details: expect.objectContaining({
					packageGraphIssues: expect.arrayContaining([
						expect.objectContaining({
							code: 'package_relationship_target',
							relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
							relationshipId: 'rIdThreaded',
						}),
					]),
					verifyIssues: expect.arrayContaining([
						expect.objectContaining({
							rule: 'threaded-comment-integrity',
							details: expect.objectContaining({
								kind: 'threaded-comment-relationship-missing-target',
							}),
						}),
					]),
				}),
			}),
		)
	})

	test('plans group orphan legacy comment relationship sidecars by comment source path', async () => {
		const input = join(TEMP_DIR, 'legacy-comment-orphan-sidecar.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeOrphanCommentRelationshipSidecarXlsx('legacy'))

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] },
		])

		expect(plan.packageGraphAudit.issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_source',
				sourcePartPath: 'xl/comments1.xml',
				relationshipPartPath: 'xl/_rels/comments1.xml.rels',
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'legacy-comment-preservation-risk',
				severity: 'warning',
				partPaths: expect.arrayContaining(['xl/comments1.xml']),
				details: expect.objectContaining({
					packageGraphIssues: expect.arrayContaining([
						expect.objectContaining({
							code: 'package_relationship_source',
							sourcePartPath: 'xl/comments1.xml',
						}),
					]),
				}),
			}),
		)
		expect(
			plan.writePolicy.diagnostics.some(
				(diagnostic) =>
					diagnostic.code === 'package-graph-audit-issue' &&
					diagnostic.partPaths?.includes('xl/comments1.xml'),
			),
		).toBe(false)
	})

	test('plans group orphan threaded comment relationship sidecars by threaded source path', async () => {
		const input = join(TEMP_DIR, 'threaded-comment-orphan-sidecar.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeOrphanCommentRelationshipSidecarXlsx('threaded'))

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] },
		])

		expect(plan.packageGraphAudit.issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_source',
				sourcePartPath: 'xl/threadedComments/threadedComment1.xml',
				relationshipPartPath: 'xl/threadedComments/_rels/threadedComment1.xml.rels',
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'threaded-comment-preservation-risk',
				severity: 'warning',
				partPaths: ['xl/threadedComments/threadedComment1.xml'],
				details: expect.objectContaining({
					packageGraphIssues: expect.arrayContaining([
						expect.objectContaining({
							code: 'package_relationship_source',
							sourcePartPath: 'xl/threadedComments/threadedComment1.xml',
						}),
					]),
				}),
			}),
		)
		expect(
			plan.writePolicy.diagnostics.some(
				(diagnostic) =>
					diagnostic.code === 'package-graph-audit-issue' &&
					diagnostic.partPaths?.includes('xl/threadedComments/threadedComment1.xml'),
			),
		).toBe(false)
	})

	test('plans attach table package graph relationship failures to table risk', async () => {
		const input = join(TEMP_DIR, 'table-missing-target.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeMissingTableTargetXlsx())

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] },
		])

		expect(plan.packageGraphAudit.issues).toContainEqual(
			expect.objectContaining({
				code: 'package_relationship_target',
				relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
				relationshipId: 'rIdTable',
				featureFamily: 'preservedTable',
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'table-preservation-risk',
				severity: 'warning',
				partPaths: expect.arrayContaining(['xl/tables/missing.xml']),
				details: expect.objectContaining({
					packageGraphIssues: expect.arrayContaining([
						expect.objectContaining({
							code: 'package_relationship_target',
							relationshipPartPath: 'xl/worksheets/_rels/sheet1.xml.rels',
							relationshipId: 'rIdTable',
						}),
					]),
					preconditions: expect.arrayContaining(['workbook-unique table names']),
				}),
			}),
		)
		expect(
			plan.writePolicy.diagnostics.some(
				(diagnostic) =>
					diagnostic.code === 'package-graph-audit-issue' &&
					diagnostic.partPaths?.includes('xl/tables/missing.xml'),
			),
		).toBe(false)
	})

	test('plans attach stale legacy comment content type overrides to comment risk', async () => {
		const input = join(TEMP_DIR, 'legacy-comment-stale-content-type.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(
			input,
			makeStaleContentTypeOverrideXlsx(
				'xl/comments1.xml',
				'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
			),
		)

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] },
		])

		expect(plan.check.issues).toContainEqual(
			expect.objectContaining({
				rule: 'package-graph-integrity',
				refs: ['[Content_Types].xml', 'xl/comments1.xml'],
				details: expect.objectContaining({
					code: 'package_content_type_override_target',
					partPath: 'xl/comments1.xml',
				}),
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'legacy-comment-preservation-risk',
				partPaths: ['xl/comments1.xml'],
				details: expect.objectContaining({
					verifyIssues: expect.arrayContaining([
						expect.objectContaining({
							rule: 'package-graph-integrity',
							details: expect.objectContaining({
								code: 'package_content_type_override_target',
								partPath: 'xl/comments1.xml',
							}),
						}),
					]),
				}),
			}),
		)
	})

	test('plans attach stale threaded comment content type overrides to threaded risk', async () => {
		const input = join(TEMP_DIR, 'threaded-comment-stale-content-type.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(
			input,
			makeStaleContentTypeOverrideXlsx(
				'xl/threadedComments/threadedComment1.xml',
				'application/vnd.ms-excel.threadedcomments+xml',
			),
		)

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] },
		])

		expect(plan.check.issues).toContainEqual(
			expect.objectContaining({
				rule: 'package-graph-integrity',
				refs: ['[Content_Types].xml', 'xl/threadedComments/threadedComment1.xml'],
				details: expect.objectContaining({
					code: 'package_content_type_override_target',
					partPath: 'xl/threadedComments/threadedComment1.xml',
				}),
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'threaded-comment-preservation-risk',
				partPaths: ['xl/threadedComments/threadedComment1.xml'],
				details: expect.objectContaining({
					verifyIssues: expect.arrayContaining([
						expect.objectContaining({
							rule: 'package-graph-integrity',
							details: expect.objectContaining({
								code: 'package_content_type_override_target',
								partPath: 'xl/threadedComments/threadedComment1.xml',
							}),
						}),
					]),
				}),
			}),
		)
	})

	test('plans attach raw legacy note VML inventory issues to comment risk', async () => {
		const input = join(TEMP_DIR, 'raw-legacy-note-vml.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		await Bun.write(input, makeRawLegacyNoteVmlXlsx())

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 'ok' }] },
		])

		expect(plan.check.issues).toContainEqual(
			expect.objectContaining({
				rule: 'legacy-comment-drawing-integrity',
				refs: ['xl/drawings/vmlDrawing1.vml'],
				details: expect.objectContaining({
					kind: 'legacy-comment-vml-without-comments-part',
					noteShapeCount: 2,
				}),
			}),
		)
		expect(plan.writePolicy.summary.commentIntegrityIssues).toBe(2)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'legacy-comment-preservation-risk',
				partPaths: expect.arrayContaining(['xl/drawings/vmlDrawing1.vml']),
				packageParts: expect.arrayContaining([
					expect.objectContaining({
						partPath: 'xl/drawings/vmlDrawing1.vml',
						featureFamily: 'preservedVml',
					}),
				]),
				details: expect.objectContaining({
					verifyIssues: expect.arrayContaining([
						expect.objectContaining({
							rule: 'legacy-comment-drawing-integrity',
							details: expect.objectContaining({
								kind: 'legacy-comment-vml-without-comments-part',
							}),
						}),
						expect.objectContaining({
							rule: 'legacy-comment-drawing-integrity',
							details: expect.objectContaining({
								kind: 'duplicate-raw-legacy-comment-vml-shape-id',
							}),
						}),
					]),
				}),
			}),
		)
	})

	test('plans explain preserved x14 conditional formatting payloads without warning on unrelated edits', async () => {
		const input = join(TEMP_DIR, 'x14-conditional-format.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().sheets[0]?.x14ConditionalFormats.push({
			index: 0,
			sqref: 'A1:A5',
			type: 'dataBar',
			priority: 4,
			formulas: [],
			preservedRuleAttributes: {
				activePresent: '1',
				'xr:uid': '{CF-UID}',
			},
			preservedRuleChildXml: [
				'<x14:extLst><x14:ext uri="{cf-extension}"><x14ac:metadata flag="1"/></x14:ext></x14:extLst>',
			],
		})
		await wb.save(input)

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 7 }] },
		])

		expect(plan.writePolicy.summary.x14ConditionalFormatExtensionPayloads).toBe(1)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'conditional-format-extension-preservation',
				severity: 'info',
				partPaths: ['xl/worksheets/sheet1.xml'],
				featureFamily: 'x14ConditionalFormatting',
				preservationPolicy: 'generated',
				details: {
					provenance: 'worksheet-extLst',
					preservationMode: 'opaque-payload-preserved-in-generated-worksheet-xml',
					semanticEditRisk: false,
					relatedOperations: [],
					x14ConditionalFormats: [
						{
							sheetName: 'Sheet1',
							sheetPartPath: 'xl/worksheets/sheet1.xml',
							source: 'x14:conditionalFormatting',
							sqref: 'A1:A5',
							index: 0,
							priority: 4,
							type: 'dataBar',
							preservedAttributeNames: ['activePresent', 'xr:uid'],
							preservedChildElements: ['x14:extLst'],
						},
					],
				},
			}),
		)
		expect(
			plan.writePolicy.diagnostics.find(
				(diagnostic) => diagnostic.code === 'conditional-format-extension-preservation',
			)?.message,
		).toContain('no conditional-format semantic edit is planned')
		expect(plan.trace.phases.find((phase) => phase.phase === 'write-policy')?.status).toBe('ok')

		const semanticEdit = await createAgentPlan(input, [
			{
				op: 'setConditionalFormat',
				sheet: 'Sheet1',
				range: 'A2:A4',
				rule: { type: 'cellIs', operator: 'greaterThan', formula: '0' },
			},
		])
		expect(semanticEdit.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'conditional-format-extension-preservation',
				severity: 'warning',
				message: expect.stringContaining('planned conditional-format semantic edits'),
				suggestedAction: expect.stringContaining('index, priority, and sqref'),
				details: expect.objectContaining({
					semanticEditRisk: true,
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'setConditionalFormat',
							sheetName: 'Sheet1',
							range: 'A2:A4',
							affectedPayloads: [
								expect.objectContaining({
									sheetName: 'Sheet1',
									sheetPartPath: 'xl/worksheets/sheet1.xml',
									source: 'x14:conditionalFormatting',
									sqref: 'A1:A5',
									index: 0,
									priority: 4,
								}),
							],
						}),
					],
				}),
			}),
		)
	})

	test('plans explain preserved x14 data validation payloads without warning on unrelated edits', async () => {
		const input = join(TEMP_DIR, 'x14-data-validation.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().sheets[0]?.x14DataValidations.push({
			index: 0,
			sqref: 'C2:C5',
			type: 'list',
			operator: 'between',
			formula1: '$A$1:$A$4',
			preservedAttributes: {
				'xr:uid': '{DV-UID}',
				showErrorMessage: '1',
			},
			preservedChildXml: ['<x14ac:metadata flag="1"><x14ac:item val="keep"/></x14ac:metadata>'],
		})
		await wb.save(input)

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 7 }] },
		])

		expect(plan.check.valid).toBe(true)
		expect(plan.modelOutput.blocked).toBe(false)
		expect(plan.modelOutput.counts.checkIssues).toBe(0)
		expect(plan.writePolicy.summary.x14DataValidationExtensionPayloads).toBe(1)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'data-validation-extension-preservation',
				severity: 'info',
				partPaths: ['xl/worksheets/sheet1.xml'],
				featureFamily: 'x14DataValidation',
				preservationPolicy: 'generated',
				details: {
					provenance: 'worksheet-extLst',
					preservationMode: 'opaque-payload-preserved-in-generated-worksheet-xml',
					semanticEditRisk: false,
					relatedOperations: [],
					x14DataValidations: [
						{
							sheetName: 'Sheet1',
							sheetPartPath: 'xl/worksheets/sheet1.xml',
							source: 'x14:dataValidations',
							sqref: 'C2:C5',
							index: 0,
							type: 'list',
							operator: 'between',
							hasFormula1: true,
							hasFormula2: false,
							preservedAttributeNames: ['xr:uid'],
							preservedChildElements: ['x14ac:metadata'],
						},
					],
				},
			}),
		)
		expect(
			plan.writePolicy.diagnostics.find(
				(diagnostic) => diagnostic.code === 'data-validation-extension-preservation',
			)?.message,
		).toContain('no data-validation semantic edit is planned')
		expect(plan.modelOutput.warnings).not.toContainEqual(expect.stringContaining('write-policy:'))

		const semanticEdit = await createAgentPlan(input, [
			{
				op: 'setDataValidation',
				sheet: 'Sheet1',
				range: 'C3:C4',
				rule: { type: 'whole', formula1: '1', formula2: '10', operator: 'between' },
			},
		])
		expect(semanticEdit.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'data-validation-extension-preservation',
				severity: 'warning',
				message: expect.stringContaining('overlap planned data-validation semantic edits'),
				suggestedAction: expect.stringContaining('index, and sqref'),
				details: expect.objectContaining({
					semanticEditRisk: true,
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'setDataValidation',
							sheetName: 'Sheet1',
							range: 'C3:C4',
							affectedPayloads: [
								expect.objectContaining({
									sheetName: 'Sheet1',
									sheetPartPath: 'xl/worksheets/sheet1.xml',
									source: 'x14:dataValidations',
									sqref: 'C2:C5',
									index: 0,
								}),
							],
						}),
					],
				}),
			}),
		)
		expect(semanticEdit.modelOutput.blocked).toBe(false)
		expect(semanticEdit.modelOutput.warnings).toEqual([
			'write-policy: 1 write policy warning(s) require inspection.',
		])
		expect(semanticEdit.modelOutput.nextActions.join('\n')).toContain(
			'Inspect writePolicy.diagnostics',
		)
	})

	test('plans warn when row or column topology edits shift preserved x14 sqref metadata', async () => {
		const input = join(TEMP_DIR, 'x14-topology-risk.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.x14ConditionalFormats.push({
			index: 0,
			sqref: 'A2:A5',
			type: 'dataBar',
			priority: 1,
			formulas: [],
			preservedRuleChildXml: ['<x14:extLst><x14:ext uri="{cf-extension}"/></x14:extLst>'],
		})
		sheet?.x14DataValidations.push({
			index: 0,
			sqref: 'C2:C5',
			type: 'list',
			formula1: '$A$1:$A$4',
			preservedAttributes: {
				'xr:uid': '{DV-TOPOLOGY-UID}',
			},
			preservedChildXml: ['<x14ac:metadata flag="1"/>'],
		})
		await wb.save(input)

		const plan = await createAgentPlan(input, [
			{ op: 'insertRows', sheet: 'Sheet1', at: 1, count: 1 },
			{ op: 'deleteCols', sheet: 'Sheet1', at: 2, count: 1 },
		])

		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'conditional-format-extension-preservation',
				severity: 'warning',
				details: expect.objectContaining({
					semanticEditRisk: true,
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'insertRows',
							sheet: 'Sheet1',
							sheetName: 'Sheet1',
							range: '2:2',
							affectedPayloads: [
								expect.objectContaining({
									source: 'x14:conditionalFormatting',
								}),
							],
						}),
					],
				}),
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'data-validation-extension-preservation',
				severity: 'warning',
				details: expect.objectContaining({
					semanticEditRisk: true,
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'insertRows',
							sheet: 'Sheet1',
							range: '2:2',
						}),
						expect.objectContaining({
							operationIndex: 1,
							op: 'deleteCols',
							sheet: 'Sheet1',
							range: 'C:C',
						}),
					],
				}),
			}),
		)
	})

	test('plans warn when table rename and column rename rewrite preserved x14 formulas', async () => {
		const input = join(TEMP_DIR, 'x14-table-reference-risk.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		expect(
			wb.apply([
				{
					op: 'setCells',
					sheet: 'Sheet1',
					updates: [
						{ ref: 'A1', value: 'Qty' },
						{ ref: 'B1', value: 'Price' },
					],
				},
				{ op: 'createTable', sheet: 'Sheet1', ref: 'A1:B3', name: 'Sales', hasHeaders: true },
			]).errors,
		).toHaveLength(0)
		const sheet = wb.getWorkbookModel().sheets[0]
		sheet?.x14ConditionalFormats.push({
			index: 0,
			sqref: 'A2:A3',
			type: 'dataBar',
			priority: 1,
			formulas: ['SUM(Sales[Qty])>0'],
			dataBar: { cfvo: [{ type: 'formula', value: 'SUM(Sales[Qty])' }] },
			preservedRuleChildXml: ['<x14:extLst><x14:ext uri="{cf-extension}"/></x14:extLst>'],
		})
		sheet?.x14DataValidations.push({
			index: 0,
			sqref: 'B2:B3',
			type: 'list',
			formula1: 'SUM(Sales[Qty])',
			preservedChildXml: ['<x14ac:metadata flag="1"/>'],
		})
		await wb.save(input)

		const tableRename = await createAgentPlan(input, [
			{ op: 'renameTable', table: 'Sales', newName: 'Revenue' },
		])
		expect(tableRename.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'conditional-format-extension-preservation',
				severity: 'warning',
				details: expect.objectContaining({
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'renameTable',
							sheet: 'Sheet1',
							table: 'Sales',
							newName: 'Revenue',
						}),
					],
				}),
			}),
		)

		const columnRename = await createAgentPlan(input, [
			{ op: 'setTableColumn', table: 'Sales', column: 'Qty', newName: 'Units' },
		])
		expect(columnRename.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'data-validation-extension-preservation',
				severity: 'warning',
				details: expect.objectContaining({
					relatedOperations: [
						expect.objectContaining({
							operationIndex: 0,
							op: 'setTableColumn',
							sheet: 'Sheet1',
							table: 'Sales',
							column: 'Qty',
							newName: 'Units',
						}),
					],
				}),
			}),
		)
	})

	test('simple x14 rules do not produce extension preservation diagnostics', async () => {
		const input = join(TEMP_DIR, 'simple-x14-rules.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		wb.getWorkbookModel().sheets[0]?.x14ConditionalFormats.push({
			index: 0,
			sqref: 'D1:D5',
			type: 'dataBar',
			priority: 1,
			formulas: [],
		})
		wb.getWorkbookModel().sheets[0]?.x14DataValidations.push({
			index: 0,
			sqref: 'A1:A5',
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '10',
			showErrorMessage: true,
		})
		await wb.save(input)

		const plan = await createAgentPlan(input, [
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'B1', value: 7 }] },
		])

		expect(plan.writePolicy.summary.x14ConditionalFormatExtensionPayloads).toBe(0)
		expect(plan.writePolicy.summary.x14DataValidationExtensionPayloads).toBe(0)
		expect(
			plan.writePolicy.diagnostics.some(
				(diagnostic) =>
					diagnostic.code === 'conditional-format-extension-preservation' ||
					diagnostic.code === 'data-validation-extension-preservation',
			),
		).toBe(false)
		expect(plan.modelOutput.warnings).not.toContainEqual(expect.stringContaining('write-policy:'))
		expect(plan.modelOutput.warnings).toEqual([])
	})

	test('partial workbook views cannot produce full-fidelity write plans', async () => {
		const wb = AscendWorkbook.create()
		const bytes = wb.toBytes()
		const partial = await AscendWorkbook.open(bytes, { mode: 'values' })
		expect(partial.inspect().load.isPartial).toBe(true)
		expect(() => partial.writePlanSummary()).toThrow(
			'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
		)
	})

	test('plan and commit emit ordered workflow progress events', async () => {
		const input = join(TEMP_DIR, 'progress.xlsx')
		const output = join(TEMP_DIR, 'progress-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		await wb.save(input)
		const ops = [{ op: 'setCells' as const, sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] }]
		const planEvents: string[] = []
		const commitEvents: string[] = []

		await createAgentPlan(input, ops, {
			onProgress: (event) => {
				planEvents.push(`${event.sequence}:${event.phase}:${event.status}`)
			},
		})
		await commitAgentPlan(input, ops, {
			output,
			onProgress: (event) => {
				commitEvents.push(`${event.sequence}:${event.phase}:${event.status}`)
			},
		})

		expect(planEvents[0]).toBe('1:hash-input:started')
		expect(planEvents).toContain('3:load-workbook:started')
		expect(planEvents.some((event) => event.includes('package-graph-audit:ok'))).toBe(true)
		expect(planEvents.at(-1)).toContain('finalize:ok')
		expect(commitEvents[0]).toBe('1:hash-input:started')
		expect(commitEvents.some((event) => event.includes('apply:ok'))).toBe(true)
		expect(commitEvents.some((event) => event.includes('post-write:ok'))).toBe(true)
		expect(commitEvents.at(-1)).toContain('finalize:ok')
	})

	test('destructive operations require explicit approval ids', async () => {
		const input = join(TEMP_DIR, 'destructive.xlsx')
		const output = join(TEMP_DIR, 'destructive-out.xlsx')
		mkdirSync(TEMP_DIR, { recursive: true })
		const wb = AscendWorkbook.create()
		wb.apply([{ op: 'addSheet', name: 'Scratch' }])
		await wb.save(input)
		const ops = [{ op: 'deleteSheet' as const, sheet: 'Scratch' }]

		const plan = await createAgentPlan(input, ops)
		const approval = plan.approvals.find((entry) => entry.kind === 'destructive-operation')
		expect(approval?.id).toBe('op:0:deletesheet')
		expect(plan.trace.phases.find((phase) => phase.phase === 'approval-audit')?.status).toBe(
			'blocked',
		)

		await expect(commitAgentPlan(input, ops, { output })).rejects.toThrow(
			'Commit requires explicit approval',
		)

		const committed = await commitAgentPlan(input, ops, {
			output,
			approvals: [approval?.id ?? ''],
		})
		expect(committed.approvals[0]?.id).toBe(approval?.id)
		expect(committed.modelOutput.blocked).toBe(false)
	})
})

function makePreservedCustomXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/custom/custom1.xml" ContentType="application/custom+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/custom/custom1.xml': encode('<custom>preserve me</custom>'),
			}),
		),
	)
}

function makeThreadedCommentXlsx(options: { readonly includePersons?: boolean } = {}): Uint8Array {
	const includePersons = options.includePersons ?? true
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
  ${
		includePersons
			? '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>'
			: ''
	}
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/worksheets/_rels/sheet1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdThreaded" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`),
				...(includePersons
					? {
							'xl/persons/person.xml':
								encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="0" displayName="Ada Lovelace"/>
  <person id="1" displayName="Grace Hopper"/>
</personList>`),
						}
					: {}),
				'xl/threadedComments/threadedComment1.xml':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="0" id="tc1" dT="2024-01-01T00:00:00.000">
    <text>Please review</text>
  </threadedComment>
  <threadedComment ref="A1" personId="1" id="tc2" parentId="tc1" dT="2024-01-02T00:00:00.000">
    <text>Reviewed</text>
  </threadedComment>
</ThreadedComments>`),
			}),
		),
	)
}

function makeStaleContentTypeOverrideXlsx(partPath: string, contentType: string): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/${partPath}" ContentType="${contentType}"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
			}),
		),
	)
}

function makeRawLegacyNoteVmlXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/worksheets/_rels/sheet1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <legacyDrawing r:id="rIdVml"/>
</worksheet>`),
				'xl/drawings/vmlDrawing1.vml':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:visible">
    <v:textbox><div>First note</div></v:textbox>
    <x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData>
  </v:shape>
  <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:visible">
    <v:textbox><div>Duplicate note</div></v:textbox>
    <x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>0</x:Column></x:ClientData>
  </v:shape>
</xml>`),
			}),
		),
	)
}

function makeMissingLegacyCommentsTargetXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/worksheets/_rels/sheet1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>`),
			}),
		),
	)
}

function makeMissingThreadedCommentsTargetXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/worksheets/_rels/sheet1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdThreaded" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/missing.xml"/>
</Relationships>`),
			}),
		),
	)
}

function makeOrphanCommentRelationshipSidecarXlsx(kind: 'legacy' | 'threaded'): Uint8Array {
	const sidecarPath =
		kind === 'legacy'
			? 'xl/_rels/comments1.xml.rels'
			: 'xl/threadedComments/_rels/threadedComment1.xml.rels'
	const sidecarXml =
		kind === 'legacy'
			? `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCommentVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="drawings/vmlDrawing1.vml"/>
</Relationships>`
			: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdThreadedMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties" Target="../metadata.xml"/>
</Relationships>`
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				[sidecarPath]: encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
${sidecarXml}`),
			}),
		),
	)
}

function makeMissingTableTargetXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <tableParts count="1"><tablePart r:id="rIdTable"/></tableParts>
</worksheet>`),
				'xl/worksheets/_rels/sheet1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/missing.xml"/>
</Relationships>`),
			}),
		),
	)
}

function makeSignedXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/_xmlsignatures/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>
  <Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rIdSignatureOrigin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/>
</Relationships>`),
				'_xmlsignatures/_rels/origin.sigs.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSignature" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="sig1.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'_xmlsignatures/origin.sigs': encode(''),
				'_xmlsignatures/sig1.xml': encode(
					`<?xml version="1.0"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>`,
				),
			}),
		),
	)
}

function makeCalcChainXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdCalcChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row>
  </sheetData>
</worksheet>`),
				'xl/calcChain.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <c r="B1" i="1"/>
</calcChain>`),
			}),
		),
	)
}

function makeExternalLinkMissingBindingXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
  <externalReferences><externalReference r:id="rIdExternal"/></externalReferences>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/externalLinks/externalLink1.xml':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rIdMissing"/>
</externalLink>`),
			}),
		),
	)
}

function makeMissingExternalLinkTargetXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/missing.xml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
  <externalReferences><externalReference r:id="rIdExternal"/></externalReferences>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
			}),
		),
	)
}

function makeMacroXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
				'xl/vbaProject.bin': encode('vba-project-bytes'),
			}),
		),
	)
}

function makeAnalyticsRefreshXlsx(): Uint8Array {
	return makeXlsx({
		'[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCache1.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>
  <Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>
</Types>`,
		'_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
		'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdPivotCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rIdSlicerCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
</Relationships>`,
		'xl/workbook.xml': `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pivotCaches><pivotCache cacheId="34" r:id="rIdPivotCache"/></pivotCaches>
  <sheets><sheet name="PivotSheet" sheetId="1" r:id="rIdSheet1"/></sheets>
</workbook>`,
		'xl/worksheets/sheet1.xml': `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPivotTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`,
		'xl/pivotTables/pivotTable1.xml': `<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="34">
  <location ref="A3:C8" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="1">
    <pivotField axis="axisPage" multipleItemSelectionAllowed="1" showAll="0">
      <items count="2">
        <item x="0"/>
        <item x="1"/>
      </items>
    </pivotField>
  </pivotFields>
  <pageFields count="1"><pageField fld="0" item="0" name="Region"/></pageFields>
</pivotTableDefinition>`,
		'xl/pivotCache/pivotCacheDefinition1.xml': `<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:id="rIdRecords" recordCount="2" refreshOnLoad="1" enableRefresh="1">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:B3" sheet="Raw"/>
  </cacheSource>
  <cacheFields count="1">
    <cacheField name="Region" databaseField="1">
      <sharedItems count="2"><s v="West"/><s v="East"/></sharedItems>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>`,
		'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdRecords" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
		'xl/pivotCache/pivotCacheRecords1.xml':
			'<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>',
		'xl/slicerCaches/slicerCache1.xml': `<?xml version="1.0"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_Region" sourceName="Region">
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
  <data><tabular pivotCacheId="34"><items count="2"><i x="0" s="1"/><i x="1"/></items></tabular></data>
</slicerCacheDefinition>`,
		'xl/slicers/slicer1.xml': `<?xml version="1.0"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="Region" cache="Slicer_Region" caption="Region"/>
</slicers>`,
		'xl/timelineCaches/timelineCache1.xml': `<?xml version="1.0"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline_Order_Date" sourceName="Order Date">
  <data><tabular pivotCacheId="34"/></data>
  <pivotTables><pivotTable name="PivotTable1"/></pivotTables>
  <state filterId="7" filterPivotName="PivotTable1" filterType="dateRange" filterTabId="2" pivotCacheId="34" singleRangeFilterState="1">
    <selection startDate="2023-01-01T00:00:00" endDate="2023-12-31T00:00:00"/>
  </state>
</timelineCacheDefinition>`,
		'xl/timelines/timeline1.xml': `<?xml version="1.0"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <timeline name="Order_Date" cache="Timeline_Order_Date" caption="Order Date"/>
</timelines>`,
	})
}

function makeVisualXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>
  <Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/worksheets/_rels/sheet1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`),
				'xl/drawings/_rels/drawing1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rIdDrawing1"/>
</worksheet>`),
				'xl/drawings/drawing1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
</xdr:wsDr>`),
				'xl/media/image1.png': encode('png-bytes'),
				'xl/charts/style1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="10"/>`),
				'xl/charts/colors1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" meth="cycle"/>`),
			}),
		),
	)
}

function makeMixedDrawingXlsx(): Uint8Array {
	return createZip(
		new Map(
			Object.entries({
				'[Content_Types].xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`),
				'_rels/.rels': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
				'xl/_rels/workbook.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
				'xl/worksheets/_rels/sheet1.xml.rels':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>`),
				'xl/workbook.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/></sheets>
</workbook>`),
				'xl/worksheets/sheet1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rIdDrawing"/>
  <legacyDrawing r:id="rIdVml"/>
</worksheet>`),
				'xl/drawings/drawing1.xml': encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="2" name="Callout"/></xdr:nvSpPr>
      <xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Original</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`),
				'xl/drawings/vmlDrawing1.vml':
					encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute;visibility:visible">
    <v:textbox><div>Legacy text</div></v:textbox>
    <x:ClientData ObjectType="Button"><x:Anchor>2, 14, 3, 6, 4, 3, 8, 0</x:Anchor></x:ClientData>
  </v:shape>
</xml>`),
			}),
		),
	)
}
