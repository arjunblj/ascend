import { describe, expect, it, setDefaultTimeout } from 'bun:test'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import type {
	ChartPartInfo,
	ChartSheetInfo,
	SheetProtection,
	WorkbookProtection,
} from '@ascend/core'
import { inspectXlsxPackageGraph, readXlsx, type XlsxPackageGraph } from '@ascend/io-xlsx'
import type { CompatibilityTier, FeatureReport } from '@ascend/schema'
import {
	type ActiveContentInfo,
	AscendWorkbook,
	type PivotCacheInfo,
	type PivotTableInfo,
} from '@ascend/sdk'
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
import { summarizeOoxmlPackage } from './package-summary.ts'

setDefaultTimeout(90_000)

const CORPUS_DIR = resolve(import.meta.dir, '../../research/excel-corpus')
const MANIFEST_PATH = resolve(CORPUS_DIR, 'manifest.json')
const SAFE_EDIT_VALUE = '__ascend_feature_contract__'
const SAFE_EDIT_TIMEOUT_MS = 180_000

interface ContractCase {
	readonly corpusName: string
	readonly rootDir: string
	readonly entry: NormalizedCorpusManifestEntry
}

interface PackageSummary {
	charts: number
	structuredCharts: number
	drawings: number
	media: number
	tables: number
	comments: number
	threadedComments: number
	pivotTables: number
	pivotCaches: number
	slicers: number
	slicerCaches: number
	timelines: number
	timelineCaches: number
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
	commentLayouts: readonly CommentLayoutSummary[]
	hasWorkbookProtection: boolean
	workbookProtection: WorkbookProtection | null
	sheetProtectionCount: number
	sheetProtections: readonly SheetProtectionSummary[]
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
	sparklineGroupCount: number
	externalReferenceCount: number
	connectionPartCount: number
	activeContentCount: number
	chartSheetCount: number
	sheetImageRefCount: number | null
	sheetDrawingObjectRefCount: number | null
	imageRefsHaveAnchors: boolean
	imageRefsHaveRelationships: boolean
	drawingObjectsHaveAnchors: boolean
	drawingObjectsHaveIdentity: boolean
	hasDrawingRefs: boolean
	activeContent: readonly ActiveContentInfo[]
	charts: readonly ChartPartInfo[]
	chartSheets: readonly ChartSheetInfo[]
	pivotTables: readonly PivotTableInfo[]
	pivotCaches: readonly PivotCacheInfo[]
	slicerCaches: readonly OoxmlLinkedCacheProbe[]
	slicers: readonly OoxmlLinkedUiProbe[]
	timelineCaches: readonly OoxmlLinkedCacheProbe[]
	timelines: readonly OoxmlLinkedUiProbe[]
}

interface CommentLayoutSummary {
	readonly sheetName: string
	readonly ref: string
	readonly shapeId?: string
	readonly anchor?: readonly number[]
	readonly row?: number
	readonly column?: number
	readonly visible?: boolean
}

interface SheetProtectionSummary {
	readonly sheetName: string
	readonly protection: SheetProtection
}

interface ContractSubject {
	readonly packageGraph: XlsxPackageGraph
	readonly packageSummary: PackageSummary
	readonly packageCounts: OoxmlPackageProbe['counts']
	readonly analytics: OoxmlAnalyticsProbe
	readonly semanticSummary: SemanticSummary
	readonly compatibilityFeatures: ReadonlyMap<string, CompatibilityFeatureSummary>
}

interface CompatibilityFeatureSummary {
	readonly tier: CompatibilityTier
	readonly count: number
}

const COMPATIBILITY_TIER_RANK: Record<CompatibilityTier, number> = {
	exact: 3,
	normalized: 2,
	preserved: 1,
	unsupported: 0,
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
		files: [
			'any_sheets.xlsx',
			'issue252.xlsx',
			'issue438.xlsx',
			'pivots.xlsx',
			'picture.xlsx',
			'vba.xlsm',
			'table-multiple.xlsx',
		],
	},
	{
		corpusName: 'closedxml',
		rootDir: resolve(import.meta.dir, '../xlsx/closedxml'),
		manifestPath: resolve(import.meta.dir, '../xlsx/closedxml/manifest.ts'),
		files: [
			'Comments_AddingComments.xlsx',
			'ConditionalFormatting_CFDataBars.xlsx',
			'ImageHandling_ImageAnchors.xlsx',
			'Misc_SheetProtection.xlsx',
			'Other_ExternalLinks_WorkbookWithExternalLink.xlsx',
			'Other_Charts_PreserveCharts_inputfile.xlsx',
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
			'textbox-hyperlink.xlsx',
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
			'sheetProtection_allLocked.xlsx',
			'sheetProtection_not_protected.xlsx',
			'SimpleStrict.xlsx',
			'SimpleWithComments.xlsx',
			'StructuredReferences.xlsx',
			'WithChart.xlsx',
			'WithDrawing.xlsx',
			'workbookProtection_workbook_structure_protected.xlsx',
		],
	},
	{
		corpusName: 'sheetjs',
		rootDir: resolve(import.meta.dir, '../xlsx/sheetjs'),
		manifestPath: resolve(import.meta.dir, '../xlsx/sheetjs/manifest.ts'),
		files: [
			'../poi/AutoFilter.xlsx',
			'../poi/formula_stress_test.xlsx',
			'../poi/merge_cells.xlsx',
			'../poi/named_ranges_2011.xlsx',
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

function isChartStyleOrColorPart(partPath: string): boolean {
	return /(^|\/)charts\/(?:style|colors)\d+\.xml$/i.test(partPath)
}

function hasSemanticChartInventory(
	packageSummary: PackageSummary,
	semanticSummary: SemanticSummary,
	compatibilityFeatures: ReadonlyMap<string, CompatibilityFeatureSummary>,
): boolean {
	if (packageSummary.structuredCharts === 0 || !compatibilityFeatures.has('preservedChart')) {
		return false
	}
	if (semanticSummary.charts.length !== packageSummary.structuredCharts) return false
	const chartSheetChartPaths = new Set(
		semanticSummary.chartSheets.flatMap((chartSheet) => chartSheet.chartPartPaths),
	)
	return semanticSummary.charts.every(
		(chart) =>
			!isChartStyleOrColorPart(chart.partPath) &&
			hasChartOwner(chart, chartSheetChartPaths) &&
			chart.chartType !== undefined &&
			chart.series.some(
				(series) => series.categoryRef !== undefined || series.valueRef !== undefined,
			),
	)
}

function hasChartOwner(chart: ChartPartInfo, chartSheetChartPaths: ReadonlySet<string>): boolean {
	return chart.sheetName !== undefined || chartSheetChartPaths.has(chart.partPath)
}

function hasSemanticDrawingInventory(
	packageSummary: PackageSummary,
	semanticSummary: SemanticSummary,
	compatibilityFeatures: ReadonlyMap<string, CompatibilityFeatureSummary>,
): boolean {
	if (packageSummary.drawings === 0) return false
	if (
		packageSummary.media > 0 &&
		!hasSemanticImageInventory(packageSummary, semanticSummary, compatibilityFeatures)
	) {
		return false
	}
	if (
		(semanticSummary.sheetDrawingObjectRefCount ?? 0) > 0 &&
		(!semanticSummary.drawingObjectsHaveAnchors || !semanticSummary.drawingObjectsHaveIdentity)
	) {
		return false
	}
	return (
		semanticSummary.hasDrawingRefs ||
		(semanticSummary.sheetImageRefCount ?? 0) > 0 ||
		(semanticSummary.sheetDrawingObjectRefCount ?? 0) > 0 ||
		compatibilityFeatures.has('drawing') ||
		compatibilityFeatures.has('preservedDrawing')
	)
}

function hasSemanticImageInventory(
	packageSummary: PackageSummary,
	semanticSummary: SemanticSummary,
	compatibilityFeatures: ReadonlyMap<string, CompatibilityFeatureSummary>,
): boolean {
	if (packageSummary.media === 0) {
		return semanticSummary.imageCount > 0 || compatibilityFeatures.has('preservedMedia')
	}
	return (
		compatibilityFeatures.has('preservedMedia') &&
		((semanticSummary.sheetImageRefCount ?? 0) === 0 ||
			(semanticSummary.sheetImageRefCount !== null &&
				semanticSummary.sheetImageRefCount <= packageSummary.media &&
				semanticSummary.imageRefsHaveAnchors &&
				semanticSummary.imageRefsHaveRelationships))
	)
}

function hasLibreOfficeActiveXControlLinkage(semanticSummary: SemanticSummary): boolean {
	return semanticSummary.activeContent.some((content) => {
		const control = content.worksheetControl
		return (
			content.kind === 'activeX' &&
			content.sheetName === 'Sheet1' &&
			content.sourceRelationshipId === 'rId3' &&
			control?.shapeId === 1025 &&
			control.name === 'CheckBox1343' &&
			control.relationshipId === 'rId3' &&
			control.controlPrRelationshipId === 'rId4' &&
			control.controlPrTarget === 'xl/media/image1.emf' &&
			control.anchor?.kind === 'twoCell' &&
			control.vmlMapOcx === true &&
			control.vmlImageTarget === 'xl/media/image1.emf'
		)
	})
}

function assertKnownVisualFixtureCoverage(
	entry: NormalizedCorpusManifestEntry,
	semanticSummary: SemanticSummary,
): void {
	const expected = KNOWN_VISUAL_FIXTURE_EXPECTATIONS.get(entry.file)
	if (!expected) return
	if (expected.sheetImageRefCount !== undefined) {
		assertFeature(
			entry,
			'visual_image_inventory',
			semanticSummary.sheetImageRefCount === expected.sheetImageRefCount &&
				semanticSummary.imageRefsHaveAnchors &&
				semanticSummary.imageRefsHaveRelationships,
			`expected ${expected.sheetImageRefCount} parsed image refs with anchors and relationships`,
		)
	}
	if (expected.sheetDrawingObjectRefCount !== undefined) {
		assertFeature(
			entry,
			'visual_drawing_object_inventory',
			semanticSummary.sheetDrawingObjectRefCount === expected.sheetDrawingObjectRefCount &&
				semanticSummary.drawingObjectsHaveAnchors &&
				semanticSummary.drawingObjectsHaveIdentity,
			`expected ${expected.sheetDrawingObjectRefCount} parsed drawing objects with anchors and identity`,
		)
	}
}

function assertKnownCommentFixtureCoverage(
	entry: NormalizedCorpusManifestEntry,
	semanticSummary: SemanticSummary,
): void {
	const expected = KNOWN_COMMENT_FIXTURE_EXPECTATIONS.get(entry.file)
	if (!expected) return
	for (const item of expected.layouts) {
		const actual = semanticSummary.commentLayouts.find(
			(layout) => layout.sheetName === item.sheetName && layout.ref === item.ref,
		)
		assertFeature(
			entry,
			'comment_vml_layout',
			actual !== undefined &&
				(item.shapeId === undefined || actual.shapeId === item.shapeId) &&
				(item.visible === undefined || actual.visible === item.visible) &&
				(item.row === undefined || actual.row === item.row) &&
				(item.column === undefined || actual.column === item.column) &&
				(item.anchor === undefined ||
					JSON.stringify(actual.anchor) === JSON.stringify(item.anchor)),
			`expected VML comment layout for ${item.sheetName}!${item.ref}`,
		)
	}
}

function assertKnownProtectionFixtureCoverage(
	entry: NormalizedCorpusManifestEntry,
	semanticSummary: SemanticSummary,
): void {
	const expected = KNOWN_PROTECTION_FIXTURE_EXPECTATIONS.get(entry.file)
	if (!expected) return
	if (expected.workbookProtection !== undefined) {
		assertFeature(
			entry,
			'workbook_protection',
			JSON.stringify(semanticSummary.workbookProtection) ===
				JSON.stringify(expected.workbookProtection),
			'expected workbook protection metadata to match fixture contract',
		)
	}
	if (expected.sheetProtections !== undefined) {
		assertFeature(
			entry,
			'sheet_protection',
			JSON.stringify(semanticSummary.sheetProtections) ===
				JSON.stringify(expected.sheetProtections),
			'expected sheet protection metadata to match fixture contract',
		)
	}
}

const KNOWN_VISUAL_FIXTURE_EXPECTATIONS = new Map<
	string,
	{ readonly sheetImageRefCount?: number; readonly sheetDrawingObjectRefCount?: number }
>([
	['ImageHandling_ImageAnchors.xlsx', { sheetImageRefCount: 7 }],
	['WithDrawing.xlsx', { sheetImageRefCount: 5, sheetDrawingObjectRefCount: 1 }],
	['picture.xlsx', { sheetImageRefCount: 2 }],
	['textbox-hyperlink.xlsx', { sheetDrawingObjectRefCount: 1 }],
])

const KNOWN_COMMENT_FIXTURE_EXPECTATIONS = new Map<
	string,
	{ readonly layouts: readonly CommentLayoutSummary[] }
>([
	[
		'SimpleWithComments.xlsx',
		{
			layouts: [
				{
					sheetName: 'Sheet1',
					ref: 'B1',
					shapeId: '_x0000_s1025',
					anchor: [2, 15, 0, 2, 4, 15, 4, 8],
					row: 0,
					column: 1,
					visible: false,
				},
				{
					sheetName: 'Sheet1',
					ref: 'B3',
					shapeId: '_x0000_s1027',
					anchor: [2, 15, 1, 7, 4, 15, 5, 13],
					row: 2,
					column: 1,
					visible: true,
				},
			],
		},
	],
])

const KNOWN_PROTECTION_FIXTURE_EXPECTATIONS = new Map<
	string,
	{
		readonly workbookProtection?: WorkbookProtection | null
		readonly sheetProtections?: readonly SheetProtectionSummary[]
	}
>([
	[
		'workbookProtection_workbook_structure_protected.xlsx',
		{ workbookProtection: { lockStructure: true }, sheetProtections: [] },
	],
	['sheetProtection_not_protected.xlsx', { workbookProtection: null, sheetProtections: [] }],
	[
		'sheetProtection_allLocked.xlsx',
		{
			workbookProtection: null,
			sheetProtections: [
				{
					sheetName: 'Foglio1',
					protection: {
						sheet: true,
						objects: true,
						scenarios: true,
						selectLockedCells: true,
						selectUnlockedCells: true,
					},
				},
			],
		},
	],
	[
		'Misc_SheetProtection.xlsx',
		{
			workbookProtection: null,
			sheetProtections: [
				{
					sheetName: 'Protected No-Password',
					protection: {
						sheet: true,
						objects: true,
						scenarios: false,
						formatCells: false,
						insertColumns: false,
						deleteColumns: false,
						deleteRows: false,
					},
				},
				{
					sheetName: 'Protected Password = 123',
					protection: {
						sheet: true,
						objects: true,
						insertColumns: false,
						insertRows: false,
						password: 'CF7A',
					},
				},
			],
		},
	],
])

function summarizePackage(bytes: Uint8Array): PackageSummary {
	const summary = summarizeOoxmlPackage(bytes)
	return {
		charts: summary.families.charts,
		structuredCharts: summary.families.structuredCharts,
		drawings: summary.families.drawings,
		media: summary.families.media,
		tables: summary.families.tables,
		comments: summary.families.comments,
		threadedComments: summary.families.threadedComments,
		pivotTables: summary.families.pivotTables,
		pivotCaches: summary.families.pivotCaches,
		slicers: summary.families.slicers,
		slicerCaches: summary.families.slicerCaches,
		timelines: summary.families.timelines,
		timelineCaches: summary.families.timelineCaches,
		macros: summary.families.macros,
		customXml: summary.families.customXml,
		externalLinks: summary.families.externalLinks,
		connections: summary.families.connections,
		calcChain: summary.families.calcChain,
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
			...(entry.sourceRelationshipId !== undefined
				? { sourceRelationshipId: entry.sourceRelationshipId }
				: {}),
			...(entry.worksheetControl !== undefined ? { worksheetControl: entry.worksheetControl } : {}),
		}))
		.sort((left, right) =>
			`${left.kind}\u0000${left.partPath}`.localeCompare(`${right.kind}\u0000${right.partPath}`),
		)
}

function summarizeCompatibilityFeatures(
	features: readonly FeatureReport[],
): ReadonlyMap<string, CompatibilityFeatureSummary> {
	const summaries = new Map<string, CompatibilityFeatureSummary>()
	for (const feature of features) {
		const existing = summaries.get(feature.feature)
		if (!existing) {
			summaries.set(feature.feature, { tier: feature.tier, count: feature.count })
			continue
		}
		const tier =
			COMPATIBILITY_TIER_RANK[feature.tier] > COMPATIBILITY_TIER_RANK[existing.tier]
				? feature.tier
				: existing.tier
		summaries.set(feature.feature, { tier, count: existing.count + feature.count })
	}
	return summaries
}

async function loadContractSubject(bytes: Uint8Array): Promise<ContractSubject> {
	const raw = readXlsx(bytes)
	expectOk(raw)
	const workbook = await AscendWorkbook.open(bytes)
	const info = workbook.inspect()
	const visuals = workbook.visualInventory()
	const packageProbe = inspectOoxmlPackageFeatures(bytes)
	const imageRefs = visuals.sheets.flatMap((sheet) => sheet.imageRefs ?? [])
	const drawingObjectRefs = visuals.sheets.flatMap((sheet) => sheet.drawingObjectRefs ?? [])
	const commentLayouts = raw.value.workbook.sheets.flatMap((sheet) =>
		[...sheet.comments.entries()].flatMap(([ref, comment]) => {
			const layout = comment.legacyDrawing
			if (!layout) return []
			return [
				{
					sheetName: sheet.name,
					ref,
					...(layout.shapeId !== undefined ? { shapeId: layout.shapeId } : {}),
					...(layout.anchor !== undefined ? { anchor: layout.anchor } : {}),
					...(layout.row !== undefined ? { row: layout.row } : {}),
					...(layout.column !== undefined ? { column: layout.column } : {}),
					...(layout.visible !== undefined ? { visible: layout.visible } : {}),
				},
			]
		}),
	)
	const sheetProtections = raw.value.workbook.sheets.flatMap((sheet) =>
		sheet.protection ? [{ sheetName: sheet.name, protection: sheet.protection }] : [],
	)
	return {
		packageGraph: inspectXlsxPackageGraph(bytes),
		packageSummary: summarizePackage(bytes),
		packageCounts: packageProbe.counts,
		analytics: packageProbe.analytics,
		semanticSummary: {
			sheetCount: info.sheetCount,
			tableCount: info.sheets.reduce((sum, sheet) => sum + (sheet.tableCount ?? 0), 0),
			commentCount: info.commentCount ?? 0,
			commentLayouts,
			hasWorkbookProtection: raw.value.workbook.workbookProtection !== null,
			workbookProtection: raw.value.workbook.workbookProtection
				? { ...raw.value.workbook.workbookProtection }
				: null,
			sheetProtectionCount: sheetProtections.length,
			sheetProtections,
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
			sparklineGroupCount: info.sparklineGroupCount ?? 0,
			externalReferenceCount: info.externalReferenceCount,
			connectionPartCount: info.connectionPartCount,
			activeContentCount: info.activeContentCount,
			chartSheetCount: info.chartSheetCount,
			sheetImageRefCount: visuals.sheetImageCount,
			sheetDrawingObjectRefCount: visuals.sheetDrawingObjectCount,
			imageRefsHaveAnchors: imageRefs.every((image) => image.anchor !== undefined),
			imageRefsHaveRelationships: imageRefs.every(
				(image) =>
					image.drawingPartPath.length > 0 && image.relId.length > 0 && image.targetPath.length > 0,
			),
			drawingObjectsHaveAnchors: drawingObjectRefs.every((object) => object.anchor !== undefined),
			drawingObjectsHaveIdentity: drawingObjectRefs.every(
				(object) =>
					object.drawingPartPath.length > 0 &&
					object.id !== undefined &&
					object.kind.length > 0 &&
					(object.name !== undefined || object.text !== undefined),
			),
			hasDrawingRefs: info.sheets.some((sheet) => sheet.hasDrawingRefs ?? false),
			activeContent: normalizeActiveContent(info.activeContent),
			charts: info.charts,
			chartSheets: info.chartSheets,
			pivotTables: info.pivotTables,
			pivotCaches: info.pivotCaches,
			slicerCaches: info.slicerCaches,
			slicers: info.slicers,
			timelineCaches: info.timelineCaches,
			timelines: info.timelines,
		},
		compatibilityFeatures: summarizeCompatibilityFeatures(raw.value.report.features),
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
	assertPackageGraphReadIntegrity(entry, subject.packageGraph)
	expectManifestCount(entry, 'worksheets', [semanticSummary.sheetCount, packageCounts.worksheets])
	expectManifestCount(entry, 'charts', [packageSummary.charts, packageCounts.charts])
	expectManifestCount(entry, 'tables', [packageSummary.tables, packageCounts.tables])
	expectManifestCount(entry, 'drawings', [packageSummary.drawings, packageCounts.drawings])
	expectManifestCount(entry, 'comments', [packageSummary.comments, packageCounts.comments])
	expectManifestCount(entry, 'workbook_protection', [packageCounts.workbook_protection])
	expectManifestCount(entry, 'sheet_protection', [packageCounts.sheet_protection])
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
	assertKnownCommentFixtureCoverage(entry, semanticSummary)
	assertFeature(
		entry,
		'workbook_protection',
		!entry.features.workbook_protection || semanticSummary.hasWorkbookProtection,
	)
	assertFeature(
		entry,
		'sheet_protection',
		!entry.features.sheet_protection || semanticSummary.sheetProtectionCount > 0,
	)
	assertFeature(
		entry,
		'protection',
		!entry.features.protection ||
			semanticSummary.hasWorkbookProtection ||
			semanticSummary.sheetProtectionCount > 0,
	)
	assertKnownProtectionFixtureCoverage(entry, semanticSummary)
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
		'timelines',
		!entry.features.timelines ||
			((packageCounts.timelines > 0 || packageCounts.timeline_caches > 0) &&
				(semanticSummary.timelineCount > 0 || semanticSummary.timelineCacheCount > 0) &&
				compatibilityFeatures.has('preservedTimeline')),
	)
	assertFeature(
		entry,
		'drawings',
		!entry.features.drawings ||
			hasSemanticDrawingInventory(packageSummary, semanticSummary, compatibilityFeatures),
	)
	assertFeature(
		entry,
		'charts',
		!entry.features.charts ||
			hasSemanticChartInventory(packageSummary, semanticSummary, compatibilityFeatures),
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
					compatibilityFeatures.has('preservedControl') ||
					compatibilityFeatures.has('preservedCustomUi')) &&
				(entry.file !== 'activex_checkbox.xlsx' ||
					hasLibreOfficeActiveXControlLinkage(semanticSummary))),
	)
	assertFeature(
		entry,
		'sparklines',
		!entry.features.sparklines ||
			(packageCounts.sparklines > 0 && semanticSummary.sparklineGroupCount > 0),
	)
	assertFeature(
		entry,
		'images_or_media',
		!entry.features.images_or_media ||
			hasSemanticImageInventory(packageSummary, semanticSummary, compatibilityFeatures),
	)
	assertKnownVisualFixtureCoverage(entry, semanticSummary)
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
	assertFeature(
		entry,
		'connections',
		!entry.features.connections ||
			(packageCounts.connections > 0 && semanticSummary.connectionPartCount > 0),
	)
	assertAnalyticsReadIntegrity(entry, subject)
}

function assertManifestEditCoverage(
	entry: NormalizedCorpusManifestEntry,
	before: ContractSubject,
	after: ContractSubject,
): void {
	assertPackageGraphEditIntegrity(entry, before.packageGraph, after.packageGraph)
	expect(after.semanticSummary.sheetCount).toBe(before.semanticSummary.sheetCount)
	expect(after.packageSummary.charts).toBe(before.packageSummary.charts)
	expect(after.packageSummary.structuredCharts).toBe(before.packageSummary.structuredCharts)
	expect(after.packageSummary.drawings).toBe(before.packageSummary.drawings)
	expect(after.packageSummary.media).toBe(before.packageSummary.media)
	expect(after.packageSummary.tables).toBe(before.packageSummary.tables)
	expect(after.packageSummary.comments).toBe(before.packageSummary.comments)
	expect(after.packageSummary.threadedComments).toBe(before.packageSummary.threadedComments)
	expect(after.packageSummary.pivotTables).toBe(before.packageSummary.pivotTables)
	expect(after.packageSummary.pivotCaches).toBe(before.packageSummary.pivotCaches)
	expect(after.packageSummary.slicers).toBe(before.packageSummary.slicers)
	expect(after.packageSummary.slicerCaches).toBe(before.packageSummary.slicerCaches)
	expect(after.packageSummary.timelines).toBe(before.packageSummary.timelines)
	expect(after.packageSummary.timelineCaches).toBe(before.packageSummary.timelineCaches)
	expect(after.packageSummary.macros).toBe(before.packageSummary.macros)
	expect(after.packageSummary.customXml).toBe(before.packageSummary.customXml)
	expect(after.packageSummary.externalLinks).toBe(before.packageSummary.externalLinks)
	expect(after.packageSummary.connections).toBe(before.packageSummary.connections)

	expect(after.semanticSummary.tableCount).toBe(before.semanticSummary.tableCount)
	expect(after.semanticSummary.commentCount).toBe(before.semanticSummary.commentCount)
	expect(after.semanticSummary.hasWorkbookProtection).toBe(
		before.semanticSummary.hasWorkbookProtection,
	)
	expect(after.semanticSummary.workbookProtection).toEqual(
		before.semanticSummary.workbookProtection,
	)
	expect(after.semanticSummary.sheetProtectionCount).toBe(
		before.semanticSummary.sheetProtectionCount,
	)
	expect(after.semanticSummary.sheetProtections).toEqual(before.semanticSummary.sheetProtections)
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
	expect(after.semanticSummary.timelineCount).toBe(before.semanticSummary.timelineCount)
	expect(after.semanticSummary.timelineCacheCount).toBe(before.semanticSummary.timelineCacheCount)
	expect(after.semanticSummary.sparklineGroupCount).toBe(before.semanticSummary.sparklineGroupCount)
	expect(after.semanticSummary.externalReferenceCount).toBe(
		before.semanticSummary.externalReferenceCount,
	)
	expect(after.semanticSummary.connectionPartCount).toBe(before.semanticSummary.connectionPartCount)
	expect(after.semanticSummary.activeContentCount).toBe(before.semanticSummary.activeContentCount)
	expect(after.semanticSummary.activeContent).toEqual(before.semanticSummary.activeContent)
	expect(after.semanticSummary.chartSheetCount).toBe(before.semanticSummary.chartSheetCount)
	expect(after.semanticSummary.chartSheets).toEqual(before.semanticSummary.chartSheets)
	expect(after.semanticSummary.charts).toEqual(before.semanticSummary.charts)
	expect(after.semanticSummary.sheetImageRefCount).toBe(before.semanticSummary.sheetImageRefCount)
	expect(after.semanticSummary.sheetDrawingObjectRefCount).toBe(
		before.semanticSummary.sheetDrawingObjectRefCount,
	)
	expect(after.semanticSummary.imageRefsHaveAnchors).toBe(
		before.semanticSummary.imageRefsHaveAnchors,
	)
	expect(after.semanticSummary.imageRefsHaveRelationships).toBe(
		before.semanticSummary.imageRefsHaveRelationships,
	)
	expect(after.semanticSummary.drawingObjectsHaveAnchors).toBe(
		before.semanticSummary.drawingObjectsHaveAnchors,
	)
	expect(after.semanticSummary.drawingObjectsHaveIdentity).toBe(
		before.semanticSummary.drawingObjectsHaveIdentity,
	)
	expect(after.semanticSummary.hasDrawingRefs).toBe(before.semanticSummary.hasDrawingRefs)

	for (const [feature, beforeFeature] of before.compatibilityFeatures) {
		if (feature === 'calcChain') continue
		if (feature === 'preservedOther') continue
		if (feature === 'preservedSignature') continue
		const afterFeature = after.compatibilityFeatures.get(feature)
		assertFeature(
			entry,
			feature,
			afterFeature !== undefined,
			`lost compatibility feature "${feature}" after safe edit`,
		)
		if (!afterFeature) continue
		if (beforeFeature.tier !== 'unsupported') {
			assertFeature(
				entry,
				feature,
				afterFeature.count >= beforeFeature.count,
				`compatibility feature "${feature}" count regressed after safe edit (${beforeFeature.count} -> ${afterFeature.count})`,
			)
		}
		assertFeature(
			entry,
			feature,
			COMPATIBILITY_TIER_RANK[afterFeature.tier] >= COMPATIBILITY_TIER_RANK[beforeFeature.tier],
			`compatibility feature "${feature}" tier regressed after safe edit (${beforeFeature.tier} -> ${afterFeature.tier})`,
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

function assertPackageGraphReadIntegrity(
	entry: NormalizedCorpusManifestEntry,
	graph: XlsxPackageGraph,
): void {
	const partPaths = new Set(graph.parts.map((part) => part.path))
	const vagueParts = graph.parts.filter(
		(part) => part.featureFamily === 'preservedOther' && !isAllowedPreservedOtherPart(part.path),
	)
	assertFeature(
		entry,
		'package_feature_classification',
		vagueParts.length === 0,
		`package graph has unclassified preservedOther parts: ${vagueParts.map((part) => part.path).join(', ')}`,
	)
	for (const relationship of graph.relationships) {
		if (relationship.targetMode?.toLowerCase() === 'external') continue
		assertFeature(
			entry,
			'package_relationship_target',
			relationship.resolvedTarget !== undefined && partPaths.has(relationship.resolvedTarget),
			`relationship ${relationship.relationshipPartPath}#${relationship.id} resolves to missing target ${relationship.resolvedTarget ?? relationship.rawTarget}`,
		)
	}
}

function assertPackageGraphEditIntegrity(
	entry: NormalizedCorpusManifestEntry,
	before: XlsxPackageGraph,
	after: XlsxPackageGraph,
): void {
	expect(after.contentTypeDefaults).toEqual(before.contentTypeDefaults)
	const afterOverrides = new Set(
		preservationRelevantContentTypeOverrides(after).map(contentTypeOverrideKey),
	)
	for (const override of preservationRelevantContentTypeOverrides(before)) {
		assertFeature(
			entry,
			'package_content_type_override',
			afterOverrides.has(contentTypeOverrideKey(override)),
			`content type override disappeared after safe edit: ${override.partPath}`,
		)
	}

	const afterParts = new Map(after.parts.map((part) => [part.path, part]))
	for (const beforePart of before.parts) {
		if (beforePart.preservationPolicy === 'discard-on-recalc') continue
		if (beforePart.preservationPolicy === 'generated') continue
		if (beforePart.preservationPolicy === 'invalidate-on-edit') {
			assertFeature(
				entry,
				'package_signature_invalidation',
				afterParts.get(beforePart.path) === undefined,
				`signature part ${beforePart.path} was retained after a generated workbook mutation`,
			)
			continue
		}
		const afterPart = afterParts.get(beforePart.path)
		assertFeature(
			entry,
			'package_preserved_part',
			afterPart !== undefined,
			`preserved package part disappeared after safe edit: ${beforePart.path}`,
		)
		if (!afterPart) continue
		assertFeature(
			entry,
			'package_preserved_part_identity',
			JSON.stringify(packagePartIdentity(afterPart)) ===
				JSON.stringify(packagePartIdentity(beforePart)),
			`preserved package part identity changed after safe edit: ${beforePart.path}`,
		)
	}

	const afterRels = new Map(
		after.relationships.map((relationship) => [
			packageRelationshipIdentityKey(relationship),
			relationship,
		]),
	)
	for (const beforeRel of before.relationships) {
		if (!beforeRel.featureFamily.startsWith('preserved')) continue
		if (beforeRel.featureFamily === 'preservedCalcChain') continue
		if (beforeRel.featureFamily === 'preservedSignature') continue
		const afterRel = afterRels.get(packageRelationshipIdentityKey(beforeRel))
		assertFeature(
			entry,
			'package_preserved_relationship',
			afterRel !== undefined,
			`preserved relationship disappeared after safe edit: ${beforeRel.relationshipPartPath}#${beforeRel.id}`,
		)
		if (!afterRel) continue
		assertFeature(
			entry,
			'package_preserved_relationship_identity',
			JSON.stringify(afterRel) === JSON.stringify(beforeRel),
			`preserved relationship identity changed after safe edit: ${beforeRel.relationshipPartPath}#${beforeRel.id}`,
		)
	}
}

function preservationRelevantContentTypeOverrides(
	graph: XlsxPackageGraph,
): readonly XlsxPackageGraph['contentTypeOverrides'][number][] {
	return graph.contentTypeOverrides.filter(
		(override) =>
			classifyPackageGraphOverrideFamily(graph, override.partPath) !== 'preservedSignature',
	)
}

function classifyPackageGraphOverrideFamily(graph: XlsxPackageGraph, partPath: string): string {
	return graph.parts.find((part) => part.path === partPath)?.featureFamily ?? 'preservedOther'
}

function contentTypeOverrideKey(
	override: XlsxPackageGraph['contentTypeOverrides'][number],
): string {
	return `${override.partPath}\u0000${override.contentType}`
}

function packagePartIdentity(part: XlsxPackageGraph['parts'][number]): Record<string, unknown> {
	return {
		contentType: part.contentType,
		contentTypeSource: part.contentTypeSource,
		ownerScope: part.ownerScope,
		sourceRelationshipPart: part.sourceRelationshipPart,
		sourceRelationshipId: part.sourceRelationshipId,
		sourceRelationshipType: part.sourceRelationshipType,
		sourceRelationshipRawTarget: part.sourceRelationshipRawTarget,
		sourceRelationshipResolvedTarget: part.sourceRelationshipResolvedTarget,
		sourceRelationshipTargetMode: part.sourceRelationshipTargetMode,
		featureFamily: part.featureFamily,
		preservationPolicy: part.preservationPolicy,
		bytePreservationExpected: part.bytePreservationExpected,
	}
}

function packageRelationshipIdentityKey(
	relationship: XlsxPackageGraph['relationships'][number],
): string {
	return `${relationship.relationshipPartPath}\u0000${relationship.id}`
}

function isAllowedPreservedOtherPart(_partPath: string): boolean {
	return false
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

			it.skipIf(!bytes)(
				'retains declared feature families after a safe edit',
				async () => {
					const sourceBytes = requireBytes(bytes)
					const before = await loadContractSubject(sourceBytes)
					const after = await applySafeEditAndReload(sourceBytes)
					assertManifestEditCoverage(entry, before, after)
				},
				SAFE_EDIT_TIMEOUT_MS,
			)
		})
	}
}
