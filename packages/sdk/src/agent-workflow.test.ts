import { afterEach, describe, expect, test } from 'bun:test'
import { existsSync, mkdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { inspectXlsxPackageGraph } from '@ascend/io-xlsx'
import { createZip, encode } from '../../io-xlsx/src/writer/zip.ts'
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
			{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'A1', value: 1 }] },
		])

		expect(plan.writePolicy.summary.externalReferences).toBe(1)
		expect(plan.writePolicy.summary.externalReferenceBindingIssues).toBe(1)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'external-link-dependency',
				severity: 'info',
				partPaths: ['xl/externalLinks/externalLink1.xml'],
				packageParts: [
					expect.objectContaining({
						partPath: 'xl/externalLinks/externalLink1.xml',
						featureFamily: 'preservedExternalLink',
						preservationPolicy: 'preserve-exact',
					}),
				],
			}),
		)
		expect(plan.writePolicy.diagnostics).toContainEqual(
			expect.objectContaining({
				code: 'external-link-binding-risk',
				severity: 'warning',
				partPaths: ['xl/externalLinks/externalLink1.xml'],
				message: expect.stringContaining('0 fallback, 1 missing'),
				suggestedAction: expect.stringContaining('rewriteExternalLink'),
			}),
		)
		expect(plan.writePolicy.ok).toBe(false)
		expect(plan.trace.phases.find((phase) => phase.phase === 'write-policy')?.status).toBe(
			'warning',
		)
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

	test('dirty cell edits preserve visual sidecars and require post-write audit inspection', async () => {
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
				severity: 'warning',
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
			}),
		)
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
