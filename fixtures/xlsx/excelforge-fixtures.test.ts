import { describe, expect, test } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
	commitAgentPlan,
	createAgentCommitPackageActionProof,
	createAgentPlan,
	inspectWorkbookOpenPlan,
} from '../../packages/sdk/src/index.ts'
import {
	normalizeManifest,
	selectManifestEntries,
	validateManifestProvenance,
} from '../corpus/manifest.ts'
import { loadManifest } from './excelforge/manifest.ts'

function loadFixture(name: string): Uint8Array {
	return readFileSync(new URL(`./excelforge/${name}`, import.meta.url))
}

function fixturePath(name: string): string {
	return fileURLToPath(new URL(`./excelforge/${name}`, import.meta.url))
}

describe('ExcelForge XLSX fixture corpus', () => {
	test('manifest has pinned provenance for the vendored MIT unknown-part fixture', async () => {
		expect(existsSync(new URL('./excelforge/source-package.json', import.meta.url))).toBe(true)
		const entries = normalizeManifest(await loadManifest())
		expect(entries).toHaveLength(1)
		expect(validateManifestProvenance(entries)).toEqual([])
		expect(selectManifestEntries(entries, { tags: ['excelforge'] })).toHaveLength(1)
		expect(selectManifestEntries(entries, { tags: ['unknown-part'] })).toHaveLength(1)
		expect(entries[0]).toMatchObject({
			file: 'Book_1_unknown_part.xlsx',
			license: 'MIT',
			sha256: '9c5426fa71ff68cc7e40e19e02b5992daf91da5754ef643d2db2f89bd70bb122',
			redistributionAllowed: true,
			vendorable: true,
		})
	})

	test('routes the real unknown package part to review before hydration', () => {
		const plan = inspectWorkbookOpenPlan(loadFixture('Book_1_unknown_part.xlsx'), {
			intent: 'edit-plan',
		})

		expect(plan).toMatchObject({
			recommendedMode: 'metadata-only',
			reviewBeforeHydration: true,
			partCount: 50,
			relationshipCount: 37,
		})
		expect(plan.riskFeatures).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					featureFamily: 'preservedOther',
					sampleParts: expect.arrayContaining(['docMetadata/LabelInfo.xml']),
				}),
			]),
		)
	})

	test('fails closed instead of silently mutating the unknown metadata part', async () => {
		const path = fixturePath('Book_1_unknown_part.xlsx')
		const dir = await mkdtemp(join(tmpdir(), 'ascend-excelforge-unknown-part-'))
		const output = join(dir, 'out.xlsx')
		const ops = [
			{ op: 'setCells', sheet: 'Projekt 1', updates: [{ ref: 'A1', value: 'probe' }] },
		] as const
		const plan = await createAgentPlan(path, ops)
		const commit = await commitAgentPlan(path, ops, {
			approvals: 'all',
			allowLoss: 'all',
			output,
		})
		const proof = createAgentCommitPackageActionProof(commit)

		expect(plan.writePolicy.ok).toBe(false)
		expect(commit.writePolicy.ok).toBe(false)
		expect(proof?.actions).toEqual(
			expect.arrayContaining([
				expect.objectContaining({
					action: 'error',
					partPath: 'docMetadata/LabelInfo.xml',
				}),
			]),
		)
	})
})
