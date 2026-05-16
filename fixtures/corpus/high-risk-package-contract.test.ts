import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { auditXlsxPackageGraphReadIntegrity, inspectXlsxPackageGraph } from '@ascend/io-xlsx'
import {
	AscendWorkbook,
	commitAgentPlan,
	createAgentCommitPackageActionProof,
	createAgentPlan,
	inspectWorkbookOpenPlan,
} from '@ascend/sdk'

function loadFixture(path: string): Uint8Array {
	return new Uint8Array(readFileSync(new URL(path, import.meta.url)))
}

function fixturePath(path: string): string {
	return fileURLToPath(new URL(path, import.meta.url))
}

describe('high-risk package corpus contract', () => {
	test('routes a public unknown package part to review and keeps mutation proof fail-closed', async () => {
		const source = loadFixture('../xlsx/excelforge/Book_1_unknown_part.xlsx')
		const graph = inspectXlsxPackageGraph(source)
		expect(graph.parts.filter((part) => part.featureFamily === 'preservedOther')).toEqual([
			expect.objectContaining({
				path: 'docMetadata/LabelInfo.xml',
				contentType: 'application/vnd.ms-office.classificationlabels+xml',
				ownerScope: 'unknown',
				featureFamily: 'preservedOther',
				preservationPolicy: 'unknown-review-required',
				bytePreservationExpected: false,
			}),
		])
		expect(auditXlsxPackageGraphReadIntegrity(graph)).toEqual([
			expect.objectContaining({
				code: 'package_feature_classification',
				severity: 'warning',
				partPath: 'docMetadata/LabelInfo.xml',
				preservationPolicy: 'unknown-review-required',
				preservationMode: 'review-required',
			}),
		])

		const openPlan = inspectWorkbookOpenPlan(source, { intent: 'edit-plan' })
		expect(openPlan).toMatchObject({
			recommendedMode: 'metadata-only',
			reviewBeforeHydration: true,
			partCount: 50,
			relationshipCount: 37,
			riskFeatures: [
				expect.objectContaining({
					featureFamily: 'preservedOther',
					category: 'unknown',
					sampleParts: ['docMetadata/LabelInfo.xml'],
				}),
			],
		})
		const metadataOnly = await AscendWorkbook.open(source, { mode: 'metadata-only' })
		expect(metadataOnly.inspect().load).toMatchObject({
			mode: 'metadata-only',
			isPartial: true,
			cellsHydrated: false,
			richSheetMetadataHydrated: false,
		})
		expect(() => metadataOnly.toBytes()).toThrow('Cannot export a partial workbook view')

		const input = fixturePath('../xlsx/excelforge/Book_1_unknown_part.xlsx')
		const dir = await mkdtemp(join(tmpdir(), 'ascend-high-risk-package-'))
		try {
			const output = join(dir, 'unknown-out.xlsx')
			const ops = [
				{
					op: 'setCells' as const,
					sheet: 'Projekt 1',
					updates: [{ ref: 'A1', value: 'unknown-part-probe' }],
				},
			]
			const plan = await createAgentPlan(input, ops)
			expect(plan.writePolicy.ok).toBe(false)
			expect(plan.lossAudit.ok).toBe(false)
			expect(plan.packageGraphAudit.ok).toBe(false)
			expect(plan.packageGraphAudit.issues).toContainEqual(
				expect.objectContaining({
					code: 'package_feature_classification',
					partPath: 'docMetadata/LabelInfo.xml',
				}),
			)
			await expect(commitAgentPlan(input, ops, { output })).rejects.toThrow(
				'Commit requires explicit approval',
			)

			const committed = await commitAgentPlan(input, ops, { output, allowLoss: 'all' })
			const proof = createAgentCommitPackageActionProof(committed)
			expect(committed.writePolicy.ok).toBe(false)
			expect(committed.postWrite.packageGraphAudit.ok).toBe(false)
			expect(proof?.actions).toEqual(
				expect.arrayContaining([
					expect.objectContaining({
						action: 'passthrough',
						partPath: 'docMetadata/LabelInfo.xml',
						bytesEqual: true,
						diagnosticCodes: expect.arrayContaining([
							'copied-through-package-parts',
							'package-graph-audit-issue',
						]),
					}),
					expect.objectContaining({
						action: 'error',
						partPath: 'docMetadata/LabelInfo.xml',
						bytesEqual: true,
						auditIssueCodes: ['package_feature_classification'],
					}),
				]),
			)
		} finally {
			await rm(dir, { recursive: true, force: true })
		}
	})

	test('requires the public encrypted fixture password and never echoes it in plans', async () => {
		const source = loadFixture('../xlsx/calamine/pass_protected.xlsx')
		expect(() => inspectWorkbookOpenPlan(source, { intent: 'edit-plan' })).toThrow(
			'requires a password',
		)
		expect(() =>
			inspectWorkbookOpenPlan(source, { intent: 'edit-plan', password: 'wrong' }),
		).toThrow('Invalid XLSX password')

		const plan = inspectWorkbookOpenPlan(source, { intent: 'edit-plan', password: '123' })
		expect(plan).toMatchObject({
			recommendedMode: 'full',
			reviewBeforeHydration: false,
			worksheetPartCount: 1,
		})
		expect(JSON.stringify(plan)).not.toContain('123')

		const workbook = await AscendWorkbook.open(source, { password: '123' })
		expect(workbook.inspect().load).toMatchObject({
			mode: 'full',
			isPartial: false,
			cellsHydrated: true,
			richSheetMetadataHydrated: true,
		})
		const changed = workbook.apply(
			[{ op: 'setCells', sheet: 'Sheet1', updates: [{ ref: 'Z10', value: 'decrypted-edit' }] }],
			{ journal: true },
		)
		expect(changed.errors).toEqual([])
		expect(changed.journal).toMatchObject({
			supported: true,
			exact: false,
			issues: [expect.objectContaining({ code: 'LOSSY_INVERSE', surface: 'package-parts' })],
		})

		const saved = workbook.toBytes()
		expect(Array.from(saved.slice(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04])
		const reopened = await AscendWorkbook.open(saved)
		expect(reopened.sheet('Sheet1')?.cell('Z10')?.value).toEqual({
			kind: 'string',
			value: 'decrypted-edit',
		})
	})
})
