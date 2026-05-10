import { describe, expect, it, setDefaultTimeout } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { readXlsx } from '@ascend/io-xlsx'
import {
	type ActiveContentInfo,
	AscendWorkbook,
	type PivotCacheInfo,
	type PivotTableInfo,
} from '@ascend/sdk'
import { extractZip } from '../../packages/io-xlsx/src/reader/zip.ts'
import type { CorpusManifestEntry, NormalizedCorpusManifestEntry } from './manifest.ts'
import { loadCorpusManifestEntries, normalizeManifest } from './manifest.ts'
import {
	inspectOoxmlPackageFeatures,
	type OoxmlAnalyticsProbe,
	type OoxmlLinkedCacheProbe,
	type OoxmlLinkedUiProbe,
	type OoxmlPackageProbe,
	type OoxmlRelationshipProbe,
} from './ooxml-feature-probe.ts'

setDefaultTimeout(90_000)

const CORPUS_DIR = resolve(import.meta.dir, '../../research/excel-corpus')
const MANIFEST_PATH = resolve(CORPUS_DIR, 'manifest.json')
const SAFE_EDIT_VALUE = '__ascend_feature_contract__'

interface ContractCase {
	readonly corpusName: string
	readonly rootDir: string
	readonly entry: NormalizedCorpusManifestEntry
}

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
	timelineCount: number
	timelineCacheCount: number
	externalReferenceCount: number
	activeContentCount: number
	hasDrawingRefs: boolean
	activeContent: readonly ActiveContentInfo[]
	pivotTables: readonly PivotTableInfo[]
	pivotCaches: readonly PivotCacheInfo[]
	slicerCaches: readonly OoxmlLinkedCacheProbe[]
	slicers: readonly OoxmlLinkedUiProbe[]
	timelineCaches: readonly OoxmlLinkedCacheProbe[]
	timelines: readonly OoxmlLinkedUiProbe[]
}

interface ContractSubject {
	readonly packageSummary: PackageSummary
	readonly packageCounts: OoxmlPackageProbe['counts']
	readonly analytics: OoxmlAnalyticsProbe
	readonly semanticSummary: SemanticSummary
	readonly compatibilityFeatures: ReadonlySet<string>
}

const VENDORED_CONTRACT_FIXTURES: readonly {
	readonly corpusName: string
	readonly rootDir: string
	readonly manifestPath: string
	readonly files: readonly string[]
}[] = [
	{
		corpusName: 'calamine',
		rootDir: resolve(import.meta.dir, '../xlsx/calamine'),
		manifestPath: resolve(import.meta.dir, '../xlsx/calamine/manifest.ts'),
		files: ['pivots.xlsx', 'picture.xlsx', 'vba.xlsm', 'table-multiple.xlsx'],
	},
	{
		corpusName: 'closedxml',
		rootDir: resolve(import.meta.dir, '../xlsx/closedxml'),
		manifestPath: resolve(import.meta.dir, '../xlsx/closedxml/manifest.ts'),
		files: [
			'Comments_AddingComments.xlsx',
			'ConditionalFormatting_CFDataBars.xlsx',
			'ImageHandling_ImageAnchors.xlsx',
			'Other_ExternalLinks_WorkbookWithExternalLink.xlsx',
			'Other_PivotTableReferenceFiles_ChartsheetAndPivotTable.xlsx',
			'Sparklines_SampleSparklines.xlsx',
			'Tables_UsingTables.xlsx',
		],
	},
	{
		corpusName: 'exceljs',
		rootDir: resolve(import.meta.dir, '../xlsx/exceljs'),
		manifestPath: resolve(import.meta.dir, '../xlsx/exceljs/manifest.ts'),
		files: ['chart-sheet.xlsx', 'formulas.xlsx', 'shared_string_with_escape.xlsx'],
	},
	{
		corpusName: 'libreoffice',
		rootDir: resolve(import.meta.dir, '../xlsx/libreoffice'),
		manifestPath: resolve(import.meta.dir, '../xlsx/libreoffice/manifest.ts'),
		files: [
			'MissingPathExternal.xlsx',
			'PivotTable_CachedDefinitionAndDataInSync.xlsx',
			'Sparklines.xlsx',
			'TableStyleTest.xlsx',
			'activex_checkbox.xlsx',
			'textLengthDataValidity.xlsx',
			'universal-content-strict.xlsx',
		],
	},
	{
		corpusName: 'poi',
		rootDir: resolve(import.meta.dir, '../xlsx/poi'),
		manifestPath: resolve(import.meta.dir, '../xlsx/poi/manifest.ts'),
		files: [
			'DataValidationEvaluations.xlsx',
			'FormulaEvalTestData_Copy.xlsx',
			'NewStyleConditionalFormattings.xlsx',
			'SimpleStrict.xlsx',
			'SimpleWithComments.xlsx',
			'StructuredReferences.xlsx',
			'WithChart.xlsx',
			'WithDrawing.xlsx',
		],
	},
]

const CONTRACT_CASES = await loadContractCases()

async function loadContractCases(): Promise<readonly ContractCase[]> {
	const cases: ContractCase[] = []
	if (existsSync(MANIFEST_PATH)) {
		const entries = normalizeManifest(
			JSON.parse(readFileSync(MANIFEST_PATH, 'utf-8')) as CorpusManifestEntry[],
		)
		for (const entry of entries) cases.push({ corpusName: 'external', rootDir: CORPUS_DIR, entry })
	}

	for (const corpus of VENDORED_CONTRACT_FIXTURES) {
		if (!existsSync(corpus.manifestPath)) continue
		const entries = normalizeManifest(await loadCorpusManifestEntries(corpus.manifestPath))
		const byFile = new Map(entries.map((entry) => [entry.file, entry]))
		for (const file of corpus.files) {
			if (!existsSync(resolve(corpus.rootDir, file))) continue
			const entry = byFile.get(file)
			if (!entry) throw new Error(`${corpus.corpusName}: missing contract fixture ${file}`)
			cases.push({ corpusName: corpus.corpusName, rootDir: corpus.rootDir, entry })
		}
	}

	return cases
}

function loadCorpusFile(rootDir: string, filename: string): Uint8Array | null {
	const path = resolve(rootDir, filename)
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

function normalizeActiveContent(entries: readonly ActiveContentInfo[]): ActiveContentInfo[] {
	return entries
		.map((entry) => ({
			kind: entry.kind,
			partPath: entry.partPath,
			contentType: entry.contentType,
			anchor: entry.anchor,
			relationshipCount: entry.relationshipCount,
			...(entry.sheetName !== undefined ? { sheetName: entry.sheetName } : {}),
			...(entry.relType !== undefined ? { relType: entry.relType } : {}),
			...(entry.byteSize !== undefined ? { byteSize: entry.byteSize } : {}),
			...(entry.opaque !== undefined ? { opaque: entry.opaque } : {}),
			...(entry.executionPolicy !== undefined ? { executionPolicy: entry.executionPolicy } : {}),
		}))
		.sort((left, right) =>
			`${left.kind}\u0000${left.partPath}`.localeCompare(`${right.kind}\u0000${right.partPath}`),
		)
}

async function loadContractSubject(bytes: Uint8Array): Promise<ContractSubject> {
	const raw = readXlsx(bytes)
	expectOk(raw)
	const workbook = await AscendWorkbook.open(bytes)
	const info = workbook.inspect()
	const packageProbe = inspectOoxmlPackageFeatures(bytes)
	return {
		packageSummary: summarizePackage(bytes),
		packageCounts: packageProbe.counts,
		analytics: packageProbe.analytics,
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
			timelineCount: info.timelineCount,
			timelineCacheCount: info.timelineCacheCount,
			externalReferenceCount: info.externalReferenceCount,
			activeContentCount: info.activeContentCount,
			hasDrawingRefs: info.sheets.some((sheet) => sheet.hasDrawingRefs ?? false),
			activeContent: normalizeActiveContent(info.activeContent),
			pivotTables: info.pivotTables,
			pivotCaches: info.pivotCaches,
			slicerCaches: info.slicerCaches,
			slicers: info.slicers,
			timelineCaches: info.timelineCaches,
			timelines: info.timelines,
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
	const { packageSummary, packageCounts, semanticSummary, compatibilityFeatures } = subject
	expectManifestCount(entry, 'worksheets', [semanticSummary.sheetCount, packageCounts.worksheets])
	expectManifestCount(entry, 'charts', [packageSummary.charts, packageCounts.charts])
	expectManifestCount(entry, 'tables', [packageSummary.tables, packageCounts.tables])
	expectManifestCount(entry, 'drawings', [packageSummary.drawings, packageCounts.drawings])
	expectManifestCount(entry, 'comments', [packageSummary.comments, packageCounts.comments])
	if (entry.features.pivot_tables) {
		expectManifestCount(entry, 'pivot_tables', [
			packageSummary.pivotTables,
			packageCounts.pivot_tables,
		])
		expectManifestCount(entry, 'pivot_caches', [
			packageSummary.pivotCaches,
			packageCounts.pivot_caches,
		])
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
			(semanticSummary.pivotTableCount > 0 && compatibilityFeatures.has('preservedPivot')),
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
			(packageSummary.charts > 0 && compatibilityFeatures.has('preservedChart')),
	)
	assertFeature(
		entry,
		'macros',
		!entry.features.macros ||
			(packageSummary.macros > 0 &&
				semanticSummary.activeContentCount > 0 &&
				semanticSummary.activeContent.some(
					(content) =>
						content.kind === 'vbaProject' &&
						content.opaque === true &&
						content.executionPolicy === 'blocked',
				) &&
				compatibilityFeatures.has('preservedMacro')),
	)
	assertFeature(
		entry,
		'active_content',
		!entry.features.active_content ||
			(packageCounts.active_content > 0 &&
				semanticSummary.activeContentCount > 0 &&
				(compatibilityFeatures.has('preservedActiveX') ||
					compatibilityFeatures.has('preservedControl'))),
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
	assertAnalyticsReadIntegrity(entry, subject)
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
	expect(after.semanticSummary.activeContentCount).toBe(before.semanticSummary.activeContentCount)
	expect(after.semanticSummary.activeContent).toEqual(before.semanticSummary.activeContent)
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
	assertAnalyticsEditIntegrity(entry, before, after)
}

function expectManifestCount(
	entry: NormalizedCorpusManifestEntry,
	countName: string,
	actuals: readonly number[],
): void {
	const expected = entry.counts[countName]
	if (expected === undefined) return
	if (actuals.includes(expected)) return
	throw new Error(
		`${entry.file}: manifest count "${countName}" expected ${expected}, observed ${actuals.join(' or ')}`,
	)
}

function assertAnalyticsReadIntegrity(
	entry: NormalizedCorpusManifestEntry,
	subject: ContractSubject,
): void {
	const { analytics, semanticSummary } = subject
	const semanticPivotPaths = new Set(semanticSummary.pivotTables.map((pivot) => pivot.partPath))
	const semanticPivotNames = new Set(
		semanticSummary.pivotTables.map((pivot) => pivot.name).filter(isDefined),
	)
	const semanticCachePaths = new Set(semanticSummary.pivotCaches.map((cache) => cache.partPath))
	const semanticCacheIds = new Set(
		semanticSummary.pivotCaches.map((cache) => cache.cacheId).filter(isDefined),
	)

	if (analytics.pivotTables.length > 0) {
		expect(semanticSummary.pivotTableCount).toBe(analytics.pivotTables.length)
	}
	if (analytics.pivotCaches.length > 0) {
		expect(semanticSummary.pivotCacheCount).toBe(analytics.pivotCaches.length)
	}
	assertRelationshipTargets(
		entry,
		'pivot table relationship',
		analytics.pivotTableRelationships,
		new Set(analytics.pivotTables.map((pivot) => pivot.partPath)),
		semanticPivotPaths,
	)
	assertRelationshipTargets(
		entry,
		'pivot cache relationship',
		analytics.pivotCacheRelationships.filter(
			(relationship) => relationship.sourcePartPath === 'xl/workbook.xml',
		),
		new Set(analytics.pivotCaches.map((cache) => cache.partPath)),
		semanticCachePaths,
	)

	for (const cache of analytics.workbookPivotCaches) {
		const relationship = analytics.pivotCacheRelationships.find(
			(candidate) => candidate.sourcePartPath === 'xl/workbook.xml' && candidate.id === cache.relId,
		)
		assertFeature(
			entry,
			'pivot_cache_relationship',
			relationship?.targetPartPath !== undefined,
			`workbook pivot cache ${cache.cacheId} does not resolve relationship ${cache.relId}`,
		)
		assertFeature(
			entry,
			'pivot_cache_semantic_cache_id',
			semanticCacheIds.has(cache.cacheId),
			`workbook pivot cache ${cache.cacheId} is missing from semantic inventory`,
		)
	}

	for (const pivot of semanticSummary.pivotTables) {
		if (pivot.cacheId === undefined) continue
		assertFeature(
			entry,
			'pivot_cache_cross_link',
			semanticCacheIds.has(pivot.cacheId),
			`pivot table ${pivot.name ?? pivot.partPath} references missing cacheId ${pivot.cacheId}`,
		)
	}

	for (const cache of analytics.pivotCaches) {
		if (!cache.recordsPartPath) continue
		const semanticCache = semanticSummary.pivotCaches.find(
			(candidate) => candidate.partPath === cache.partPath,
		)
		assertFeature(
			entry,
			'pivot_cache_records',
			analytics.pivotCacheRecords.includes(cache.recordsPartPath),
			`pivot cache ${cache.partPath} points to missing records part ${cache.recordsPartPath}`,
		)
		assertFeature(
			entry,
			'pivot_cache_records_semantic',
			semanticCache?.recordsPartPath === cache.recordsPartPath,
			`pivot cache ${cache.partPath} records relationship is not surfaced semantically`,
		)
	}

	assertLinkedCacheIntegrity(entry, 'slicer', analytics.slicerCaches, analytics.slicers, {
		cacheRelationships: analytics.slicerCacheRelationships,
		semanticCaches: semanticSummary.slicerCaches,
		semanticUis: semanticSummary.slicers,
		semanticPivotNames,
	})
	assertLinkedCacheIntegrity(entry, 'timeline', analytics.timelineCaches, analytics.timelines, {
		semanticCaches: semanticSummary.timelineCaches,
		semanticUis: semanticSummary.timelines,
		semanticPivotNames,
	})
}

function assertLinkedCacheIntegrity(
	entry: NormalizedCorpusManifestEntry,
	label: 'slicer' | 'timeline',
	packageCaches: readonly OoxmlLinkedCacheProbe[],
	packageUis: readonly OoxmlLinkedUiProbe[],
	context: {
		readonly cacheRelationships?: readonly OoxmlRelationshipProbe[]
		readonly semanticCaches: readonly OoxmlLinkedCacheProbe[]
		readonly semanticUis: readonly OoxmlLinkedUiProbe[]
		readonly semanticPivotNames: ReadonlySet<string>
	},
): void {
	if (packageCaches.length > 0) expect(context.semanticCaches).toHaveLength(packageCaches.length)
	if (packageUis.length > 0) expect(context.semanticUis).toHaveLength(packageUis.length)
	const packageCachePaths = new Set(packageCaches.map((cache) => cache.partPath))
	const semanticCachePaths = new Set(context.semanticCaches.map((cache) => cache.partPath))
	const packageCacheNames = new Set(packageCaches.map((cache) => cache.name).filter(isDefined))

	if (context.cacheRelationships) {
		assertRelationshipTargets(
			entry,
			`${label} cache relationship`,
			context.cacheRelationships,
			packageCachePaths,
			semanticCachePaths,
		)
	}
	for (const cache of packageCaches) {
		assertFeature(
			entry,
			`${label}_cache_semantic_part`,
			semanticCachePaths.has(cache.partPath),
			`${label} cache ${cache.partPath} is missing from semantic inventory`,
		)
		for (const pivotTableName of cache.pivotTableNames) {
			if (context.semanticPivotNames.size === 0) continue
			assertFeature(
				entry,
				`${label}_pivot_table_cross_link`,
				context.semanticPivotNames.has(pivotTableName),
				`${label} cache ${cache.name ?? cache.partPath} references missing pivot table ${pivotTableName}`,
			)
		}
	}
	for (const ui of packageUis) {
		if (!ui.cacheName) continue
		assertFeature(
			entry,
			`${label}_cache_name_cross_link`,
			packageCacheNames.has(ui.cacheName),
			`${label} ${ui.name ?? ui.partPath} references missing cache ${ui.cacheName}`,
		)
	}
}

function assertRelationshipTargets(
	entry: NormalizedCorpusManifestEntry,
	label: string,
	relationships: readonly OoxmlRelationshipProbe[],
	packageTargets: ReadonlySet<string>,
	semanticTargets: ReadonlySet<string>,
): void {
	for (const relationship of relationships) {
		const target = relationship.targetPartPath
		assertFeature(
			entry,
			label,
			target !== undefined && packageTargets.has(target) && semanticTargets.has(target),
			`${label} ${relationship.sourcePartPath}#${relationship.id} does not resolve to a surfaced package part`,
		)
	}
}

function assertAnalyticsEditIntegrity(
	entry: NormalizedCorpusManifestEntry,
	before: ContractSubject,
	after: ContractSubject,
): void {
	expectAnalyticsPaths(
		entry,
		'pivot tables',
		before.analytics.pivotTables,
		after.analytics.pivotTables,
	)
	expectAnalyticsPaths(
		entry,
		'pivot caches',
		before.analytics.pivotCaches,
		after.analytics.pivotCaches,
	)
	expect(after.analytics.pivotCacheRecords).toEqual(before.analytics.pivotCacheRecords)
	expectAnalyticsPaths(
		entry,
		'slicer caches',
		before.analytics.slicerCaches,
		after.analytics.slicerCaches,
	)
	expectAnalyticsPaths(entry, 'slicers', before.analytics.slicers, after.analytics.slicers)
	expectAnalyticsPaths(
		entry,
		'timeline caches',
		before.analytics.timelineCaches,
		after.analytics.timelineCaches,
	)
	expectAnalyticsPaths(entry, 'timelines', before.analytics.timelines, after.analytics.timelines)
	expect(after.analytics.workbookPivotCaches).toEqual(before.analytics.workbookPivotCaches)
	expect(expectRelationshipSignature(after.analytics.pivotTableRelationships)).toEqual(
		expectRelationshipSignature(before.analytics.pivotTableRelationships),
	)
	expect(expectRelationshipSignature(after.analytics.pivotCacheRelationships)).toEqual(
		expectRelationshipSignature(before.analytics.pivotCacheRelationships),
	)
	expect(expectRelationshipSignature(after.analytics.pivotCacheRecordRelationships)).toEqual(
		expectRelationshipSignature(before.analytics.pivotCacheRecordRelationships),
	)
	expect(expectRelationshipSignature(after.analytics.slicerCacheRelationships)).toEqual(
		expectRelationshipSignature(before.analytics.slicerCacheRelationships),
	)
}

function expectAnalyticsPaths(
	entry: NormalizedCorpusManifestEntry,
	label: string,
	before: readonly { readonly partPath: string }[],
	after: readonly { readonly partPath: string }[],
): void {
	assertFeature(
		entry,
		label,
		after.map((item) => item.partPath).join('\n') ===
			before.map((item) => item.partPath).join('\n'),
		`${label} package part set changed after safe edit`,
	)
}

function expectRelationshipSignature(
	relationships: readonly OoxmlRelationshipProbe[],
): readonly string[] {
	return relationships.map(
		(relationship) =>
			`${relationship.sourcePartPath}#${relationship.id}:${relationship.type}->${relationship.targetPartPath ?? relationship.target}`,
	)
}

function isDefined<T>(value: T | undefined): value is T {
	return value !== undefined
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

if (CONTRACT_CASES.length === 0) {
	describe.skip('corpus feature contract', () => {
		it('skips when no corpus manifests are available', () => {})
	})
} else {
	for (const { corpusName, rootDir, entry } of CONTRACT_CASES) {
		describe(`corpus feature contract: ${corpusName}/${entry.file}`, () => {
			const bytes = loadCorpusFile(rootDir, entry.file)

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
