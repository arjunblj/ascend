import { describe, expect, it, setDefaultTimeout } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { readXlsx } from '@ascend/io-xlsx'
import { AscendWorkbook } from '@ascend/sdk'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'
import type { CorpusManifestEntry, NormalizedCorpusManifestEntry } from './manifest.ts'
import { normalizeManifest } from './manifest.ts'

setDefaultTimeout(45_000)

const CORPUS_DIR = resolve(import.meta.dir, '../../research/excel-corpus')
const MANIFEST_PATH = resolve(CORPUS_DIR, 'manifest.json')
const SAFE_EDIT_VALUE = '__ascend_feature_contract__'

interface PackageSummary {
	charts: number
	drawings: number
	media: number
	tables: number
	comments: number
	threadedComments: number
	pivotTables: number
	pivotCaches: number
	slicers: number
	slicerCaches: number
	macros: number
	customXml: number
	externalLinks: number
	connections: number
	calcChain: number
}

interface SemanticSummary {
	sheetCount: number
	tableCount: number
	commentCount: number
	conditionalFormatCount: number
	dataValidationCount: number
	imageCount: number
	hyperlinkCount: number
	mergeCount: number
	definedNameCount: number
	pivotTableCount: number
	pivotCacheCount: number
	slicerCount: number
	slicerCacheCount: number
	externalReferenceCount: number
	hasDrawingRefs: boolean
}

interface ContractSubject {
	readonly packageSummary: PackageSummary
	readonly semanticSummary: SemanticSummary
	readonly compatibilityFeatures: ReadonlySet<string>
}

const HAS_MANIFEST = existsSync(MANIFEST_PATH)
const MANIFEST = HAS_MANIFEST
	? normalizeManifest(JSON.parse(readFileSync(MANIFEST_PATH, 'utf-8')) as CorpusManifestEntry[])
	: []

function loadCorpusFile(filename: string): Uint8Array | null {
	const path = resolve(CORPUS_DIR, filename)
	if (!existsSync(path)) return null
	return new Uint8Array(readFileSync(path))
}

function requireBytes(bytes: Uint8Array | null): Uint8Array {
	if (!bytes) throw new Error('Corpus file not available')
	return bytes
}

function countPaths(paths: readonly string[], pattern: RegExp): number {
	return paths.filter((path) => pattern.test(path)).length
}

function summarizePackage(bytes: Uint8Array): PackageSummary {
	const archive = extractZip(bytes)
	const paths = [...archive.entries()].map((entry) => entry.path)
	return {
		charts: countPaths(paths, /^xl\/(charts|chartEx)\//),
		drawings: countPaths(paths, /^xl\/drawings\//),
		media: countPaths(paths, /^xl\/media\//),
		tables: countPaths(paths, /^xl\/tables\//),
		comments: countPaths(paths, /^xl\/comments\d+\.xml$/),
		threadedComments: countPaths(paths, /^xl\/threadedComments\//),
		pivotTables: countPaths(paths, /^xl\/pivotTables\//),
		pivotCaches: countPaths(paths, /^xl\/pivotCache/),
		slicers: countPaths(paths, /^xl\/slicers\//),
		slicerCaches: countPaths(paths, /^xl\/slicerCaches\//),
		macros: countPaths(paths, /^xl\/vbaProject/i),
		customXml: countPaths(paths, /^customXml\//),
		externalLinks: countPaths(paths, /^xl\/externalLinks\//),
		connections: countPaths(paths, /^xl\/connections\.xml$/),
		calcChain: countPaths(paths, /^xl\/calcChain\.xml$/),
	}
}

async function loadContractSubject(bytes: Uint8Array): Promise<ContractSubject> {
	const raw = readXlsx(bytes)
	expectOk(raw)
	const workbook = await AscendWorkbook.open(bytes)
	const info = workbook.inspect()
	return {
		packageSummary: summarizePackage(bytes),
		semanticSummary: {
			sheetCount: info.sheetCount,
			tableCount: info.sheets.reduce((sum, sheet) => sum + (sheet.tableCount ?? 0), 0),
			commentCount: info.commentCount ?? 0,
			conditionalFormatCount: info.conditionalFormatCount ?? 0,
			dataValidationCount: info.dataValidationCount ?? 0,
			imageCount: info.imageCount ?? 0,
			hyperlinkCount: info.sheets.reduce((sum, sheet) => sum + (sheet.hyperlinkCount ?? 0), 0),
			mergeCount: workbook.sheets.reduce(
				(sum, sheetName) => sum + (workbook.sheet(sheetName)?.merges.length ?? 0),
				0,
			),
			definedNameCount: info.definedNameDetails.length,
			pivotTableCount: info.pivotTableCount,
			pivotCacheCount: info.pivotCacheCount,
			slicerCount: info.slicerCount,
			slicerCacheCount: info.slicerCacheCount,
			externalReferenceCount: info.externalReferenceCount,
			hasDrawingRefs: info.sheets.some((sheet) => sheet.hasDrawingRefs ?? false),
		},
		compatibilityFeatures: new Set(raw.value.report.features.map((feature) => feature.feature)),
	}
}

async function applySafeEditAndReload(bytes: Uint8Array): Promise<ContractSubject> {
	const workbook = await AscendWorkbook.open(bytes)
	const probe = pickProbeTarget(workbook)
	const apply = workbook.apply([
		{
			op: 'setCells',
			sheet: probe.sheet,
			updates: [{ ref: probe.ref, value: SAFE_EDIT_VALUE }],
		},
	])
	expect(apply.errors).toHaveLength(0)
	if (apply.recalcRequired) {
		const recalc = workbook.recalc()
		expect(recalc.errors).toHaveLength(0)
	}
	return loadContractSubject(workbook.toBytes())
}

function pickProbeTarget(workbook: AscendWorkbook): { sheet: string; ref: string } {
	const rankedSheets = workbook.sheets
		.map((name) => workbook.sheet(name))
		.filter((sheet): sheet is NonNullable<typeof sheet> => sheet !== undefined)
		.sort((left, right) => scoreSheet(right) - scoreSheet(left))

	const sheet = rankedSheets[0]
	if (!sheet) throw new Error('Workbook has no sheets')
	const used = sheet.usedRange()
	if (!used) return { sheet: sheet.name, ref: 'A1' }
	return {
		sheet: sheet.name,
		ref: `${columnLabel(used.end.col + 1)}${used.end.row + 2}`,
	}
}

function scoreSheet(sheet: NonNullable<ReturnType<AscendWorkbook['sheet']>>): number {
	let score = 0
	if (sheet.state === 'visible') score += 10
	if (!sheet.protection?.sheet) score += 5
	if (!sheet.autoFilter) score += 1
	return score
}

function columnLabel(col: number): string {
	let n = col
	let label = ''
	while (n >= 0) {
		label = String.fromCharCode(65 + (n % 26)) + label
		n = Math.floor(n / 26) - 1
	}
	return label
}

function assertManifestReadCoverage(
	entry: NormalizedCorpusManifestEntry,
	subject: ContractSubject,
): void {
	const { packageSummary, semanticSummary, compatibilityFeatures } = subject
	expect(semanticSummary.sheetCount).toBe(entry.counts.worksheets)
	expect(packageSummary.charts).toBe(entry.counts.charts)
	expect(packageSummary.tables).toBe(entry.counts.tables)
	expect(packageSummary.drawings).toBe(entry.counts.drawings)
	expect(packageSummary.comments).toBe(entry.counts.comments)
	if (entry.features.pivot_tables) {
		expect(packageSummary.pivotTables).toBe(entry.counts.pivot_tables)
		expect(packageSummary.pivotCaches).toBe(entry.counts.pivot_caches)
	}

	assertFeature(entry, 'tables', !entry.features.tables || semanticSummary.tableCount > 0)
	assertFeature(entry, 'comments', !entry.features.comments || semanticSummary.commentCount > 0)
	assertFeature(
		entry,
		'threaded_comments',
		!entry.features.threaded_comments ||
			(packageSummary.threadedComments > 0 &&
				compatibilityFeatures.has('preservedThreadedComments')),
	)
	assertFeature(
		entry,
		'conditional_formatting',
		!entry.features.conditional_formatting || semanticSummary.conditionalFormatCount > 0,
	)
	assertFeature(
		entry,
		'data_validations',
		!entry.features.data_validations || semanticSummary.dataValidationCount > 0,
	)
	assertFeature(
		entry,
		'merged_cells',
		!entry.features.merged_cells || semanticSummary.mergeCount > 0,
	)
	assertFeature(
		entry,
		'hyperlinks',
		!entry.features.hyperlinks || semanticSummary.hyperlinkCount > 0,
	)
	assertFeature(
		entry,
		'defined_names',
		!entry.features.defined_names || semanticSummary.definedNameCount > 0,
	)
	assertFeature(
		entry,
		'pivot_tables',
		!entry.features.pivot_tables ||
			(semanticSummary.pivotTableCount > 0 &&
				compatibilityFeatures.has('pivotTable') &&
				compatibilityFeatures.has('preservedPivot')),
	)
	assertFeature(
		entry,
		'slicers',
		!entry.features.slicers ||
			(semanticSummary.slicerCount > 0 &&
				semanticSummary.slicerCacheCount > 0 &&
				packageSummary.slicers > 0 &&
				packageSummary.slicerCaches > 0 &&
				compatibilityFeatures.has('preservedSlicer')),
	)
	assertFeature(
		entry,
		'drawings',
		!entry.features.drawings ||
			(packageSummary.drawings > 0 &&
				(semanticSummary.hasDrawingRefs ||
					semanticSummary.imageCount > 0 ||
					compatibilityFeatures.has('drawing') ||
					compatibilityFeatures.has('preservedDrawing'))),
	)
	assertFeature(
		entry,
		'charts',
		!entry.features.charts ||
			(packageSummary.charts > 0 &&
				compatibilityFeatures.has('chart') &&
				compatibilityFeatures.has('preservedChart')),
	)
	assertFeature(
		entry,
		'macros',
		!entry.features.macros ||
			(packageSummary.macros > 0 &&
				compatibilityFeatures.has('vbaProject') &&
				compatibilityFeatures.has('preservedMacro')),
	)
	assertFeature(
		entry,
		'images_or_media',
		!entry.features.images_or_media ||
			packageSummary.media > 0 ||
			semanticSummary.imageCount > 0 ||
			compatibilityFeatures.has('preservedMedia'),
	)
	assertFeature(
		entry,
		'custom_xml',
		!entry.features.custom_xml ||
			(packageSummary.customXml > 0 && compatibilityFeatures.has('preservedCustomXml')),
	)
	assertFeature(
		entry,
		'calc_chain',
		!entry.features.calc_chain ||
			(packageSummary.calcChain > 0 && compatibilityFeatures.has('calcChain')),
	)
	assertFeature(
		entry,
		'external_links',
		!entry.features.external_links ||
			packageSummary.externalLinks > 0 ||
			semanticSummary.externalReferenceCount > 0,
	)
	assertFeature(entry, 'connections', !entry.features.connections || packageSummary.connections > 0)
}

function assertManifestEditCoverage(
	entry: NormalizedCorpusManifestEntry,
	before: ContractSubject,
	after: ContractSubject,
): void {
	expect(after.semanticSummary.sheetCount).toBe(before.semanticSummary.sheetCount)
	expect(after.packageSummary.charts).toBe(before.packageSummary.charts)
	expect(after.packageSummary.drawings).toBe(before.packageSummary.drawings)
	expect(after.packageSummary.media).toBe(before.packageSummary.media)
	expect(after.packageSummary.tables).toBe(before.packageSummary.tables)
	expect(after.packageSummary.comments).toBe(before.packageSummary.comments)
	expect(after.packageSummary.threadedComments).toBe(before.packageSummary.threadedComments)
	expect(after.packageSummary.pivotTables).toBe(before.packageSummary.pivotTables)
	expect(after.packageSummary.pivotCaches).toBe(before.packageSummary.pivotCaches)
	expect(after.packageSummary.slicers).toBe(before.packageSummary.slicers)
	expect(after.packageSummary.slicerCaches).toBe(before.packageSummary.slicerCaches)
	expect(after.packageSummary.macros).toBe(before.packageSummary.macros)
	expect(after.packageSummary.customXml).toBe(before.packageSummary.customXml)
	expect(after.packageSummary.externalLinks).toBe(before.packageSummary.externalLinks)
	expect(after.packageSummary.connections).toBe(before.packageSummary.connections)

	expect(after.semanticSummary.tableCount).toBe(before.semanticSummary.tableCount)
	expect(after.semanticSummary.commentCount).toBe(before.semanticSummary.commentCount)
	expect(after.semanticSummary.conditionalFormatCount).toBe(
		before.semanticSummary.conditionalFormatCount,
	)
	expect(after.semanticSummary.dataValidationCount).toBe(before.semanticSummary.dataValidationCount)
	expect(after.semanticSummary.imageCount).toBe(before.semanticSummary.imageCount)
	expect(after.semanticSummary.hyperlinkCount).toBe(before.semanticSummary.hyperlinkCount)
	expect(after.semanticSummary.mergeCount).toBe(before.semanticSummary.mergeCount)
	expect(after.semanticSummary.definedNameCount).toBe(before.semanticSummary.definedNameCount)
	expect(after.semanticSummary.pivotTableCount).toBe(before.semanticSummary.pivotTableCount)
	expect(after.semanticSummary.pivotCacheCount).toBe(before.semanticSummary.pivotCacheCount)
	expect(after.semanticSummary.slicerCount).toBe(before.semanticSummary.slicerCount)
	expect(after.semanticSummary.slicerCacheCount).toBe(before.semanticSummary.slicerCacheCount)
	expect(after.semanticSummary.externalReferenceCount).toBe(
		before.semanticSummary.externalReferenceCount,
	)
	expect(after.semanticSummary.hasDrawingRefs).toBe(before.semanticSummary.hasDrawingRefs)

	for (const feature of before.compatibilityFeatures) {
		if (feature === 'calcChain') continue
		if (feature === 'preservedOther') continue
		assertFeature(
			entry,
			feature,
			after.compatibilityFeatures.has(feature),
			`lost compatibility feature "${feature}" after safe edit`,
		)
	}
	if (entry.features.calc_chain) {
		assertFeature(
			entry,
			'calc_chain',
			after.compatibilityFeatures.has('calcChain') ||
				after.compatibilityFeatures.has('formulaFreshness'),
			'expected calc-chain workbooks to preserve calcChain or surface formulaFreshness after edit',
		)
	}
}

function assertFeature(
	entry: NormalizedCorpusManifestEntry,
	feature: string,
	condition: boolean,
	reason?: string,
): void {
	if (condition) return
	throw new Error(
		`${entry.file}: expected feature "${feature}" coverage${reason ? ` (${reason})` : ''}`,
	)
}

function expectOk<T, E extends { message: string }>(
	result: { ok: true; value: T } | { ok: false; error: E },
): asserts result is { ok: true; value: T } {
	expect(result.ok).toBe(true)
	if (!result.ok) throw new Error(result.error.message)
}

if (!HAS_MANIFEST) {
	describe.skip('corpus feature contract', () => {
		it('skips when the external corpus manifest is unavailable', () => {})
	})
} else {
	for (const entry of MANIFEST) {
		describe(`corpus feature contract: ${entry.file}`, () => {
			const bytes = loadCorpusFile(entry.file)

			it.skipIf(!bytes)('surfaces every declared feature family on read', async () => {
				const subject = await loadContractSubject(requireBytes(bytes))
				assertManifestReadCoverage(entry, subject)
			})

			it.skipIf(!bytes)('retains declared feature families after a safe edit', async () => {
				const sourceBytes = requireBytes(bytes)
				const before = await loadContractSubject(sourceBytes)
				const after = await applySafeEditAndReload(sourceBytes)
				assertManifestEditCoverage(entry, before, after)
			})
		})
	}
}
