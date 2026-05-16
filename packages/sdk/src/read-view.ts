import {
	type AutoFilter,
	type CellStyle,
	type ChartPartInfo,
	cloneActiveContentInfo,
	clonePivotCacheInfo,
	clonePivotTableInfo,
	cloneStyle,
	cloneX14ConditionalFormatInfo,
	cloneX14DataValidationInfo,
	indexToColumn,
	parseA1,
	parseRange,
	pivotDataFieldCaptionsMatch,
	type RangeRef,
	type SheetDrawingObjectRef,
	type SheetImageAnchor,
	type SheetImageRef,
	type Workbook,
	type WorkbookDocumentProperties,
} from '@ascend/core'
import {
	analyzeWorkbook,
	analyzeWorkbookDependencies,
	analyzeWorkbookFormulas,
	createSnapshot,
	diffWorkbooks,
	resolveCellFormulaText,
	type WorkbookAnalysis,
	type WorkbookDependencyAnalysis,
	type WorkbookDiff,
	type WorkbookFormulaAnalysis,
	type WorkbookSnapshot,
} from '@ascend/engine'
import {
	type FormulaNode,
	normalizeFormulaInput,
	parseFormula,
	printFormula,
} from '@ascend/formulas'
import type { CellUpdate, CellValue, CompatibilityReport, Operation } from '@ascend/schema'
import { EMPTY } from '@ascend/schema'
import { trace as verifyTrace } from '@ascend/verify'
import { getCapability, isCapabilityGap } from './capabilities.ts'
import { partialDependencyCheckIssue } from './check-issues.ts'
import { type FormatDisplayOptions, formatStyledDisplayCellValue } from './format-helpers.ts'
import {
	buildFormulaInfo,
	collectFormulaReferences,
	collectFunctionNames,
	flattenLegacyReferenceTexts,
	hasVolatileFunction,
	tokenizeFormulaInput,
} from './formula-info.ts'
import { type CellSelector, normalizeCellSelector } from './ref-selectors.ts'
import { SheetHandle } from './sheet-handle.ts'
import { TableHandle } from './table-handle.ts'
import type {
	ActiveContentInfo,
	AgentColumnSummary,
	AgentReadOptions,
	AgentSampleRow,
	AgentViewOptions,
	AgentViewResult,
	CapabilityWarningInfo,
	CheckIssue,
	CompactCellInfo,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	DefinedNameInfo,
	ExternalReferenceInfo,
	ExternalReferenceUsageInfo,
	FlatCellValue,
	FormulaInfo,
	FormulaReferenceInfo,
	GetPivotDataFilter,
	GetPivotDataQuery,
	GetPivotDataResult,
	PivotCacheDecodedValueInfo,
	PivotCacheInfo,
	PivotCacheMaterializedRowInfo,
	PivotCacheRecordInfo,
	PivotCacheRecordValueInfo,
	PivotCacheRecordValueKindCount,
	PivotCacheRowsOptions,
	PivotFieldInfo,
	PivotFieldItemInfo,
	PivotFieldReference,
	PivotOutputAuditInfo,
	PivotOutputAuditMismatchInfo,
	PivotOutputMaterializeMode,
	PivotOutputMaterializeOpsResult,
	PivotOutputMaterializeOptions,
	PivotOutputMaterializeUnsupportedInfo,
	PivotRefreshPlanInfo,
	PivotRefreshRecommendedOp,
	PivotTableInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
	ReadSnapshotValidationResult,
	SheetInfo,
	SheetInspectInfo,
	SlicerCacheInfo,
	SlicerInfo,
	TableInfo,
	TimelineCacheInfo,
	TimelineInfo,
	TraceResult,
	WorkbookConnectionPartInfo,
	WorkbookDataModelPartInfo,
	WorkbookInfo,
	WorkbookLoadInfo,
	WorkbookReadSnapshotInfo,
	WorkbookRefreshMetadataEntry,
	WorkbookRefreshMetadataInfo,
	WorkbookVisualInventoryInfo,
} from './types.ts'
import {
	buildWorkbookTrustReport,
	type WorkbookTrustReport,
	type WorkbookTrustReportOptions,
} from './workbook-trust.ts'

/**
 * Read-only view over a workbook. Caching strategy:
 * - workbookInfoCache: single cached WorkbookInfo from inspect()
 * - sheetInspectCache: per-sheet SheetInspectInfo
 * - formulaInfoCache: per-cell-ref FormulaInfo (formula() results)
 * - sheetHandleCache / tableHandleCache: handles reused for repeated access
 * - formulaAnalysis/dependencyAnalysis: cached at engine level (WeakMap keyed by workbook)
 */
export class WorkbookReadView {
	protected wb: Workbook
	protected compat: CompatibilityReport
	protected loadInfo: WorkbookLoadInfo
	private workbookInfoCache: WorkbookInfo | undefined
	private readonly sheetInspectCache = new Map<string, SheetInspectInfo | undefined>()
	private readonly formulaInfoCache = new Map<string, FormulaInfo | undefined>()
	private readonly sheetHandleCache = new Map<string, SheetHandle>()
	private readonly tableHandleCache = new Map<string, TableHandle>()

	constructor(workbook: Workbook, report: CompatibilityReport, loadInfo: WorkbookLoadInfo) {
		this.wb = workbook
		this.compat = report
		this.loadInfo = loadInfo
	}

	getWorkbookModel(): Workbook {
		return this.wb
	}

	readSnapshotInfo(): WorkbookReadSnapshotInfo {
		return this.currentReadSnapshot()
	}

	validateReadSnapshot(snapshot: WorkbookReadSnapshotInfo | string): ReadSnapshotValidationResult {
		const expected = this.currentReadSnapshot()
		const receivedToken = typeof snapshot === 'string' ? snapshot : snapshot.token
		if (receivedToken === expected.token) {
			return { ok: true, current: true, expected, receivedToken }
		}
		return {
			ok: false,
			current: false,
			expected,
			receivedToken,
			error: {
				code: 'VALIDATION_ERROR',
				message:
					'Read snapshot token is stale; refresh the viewport before applying derived patches.',
				retryable: false,
				details: {
					rule: 'stale-read-snapshot',
					expectedToken: expected.token,
					receivedToken,
					expectedGenerations: expected.generations,
					load: expected.load,
				},
				suggestedFix:
					'Read the viewport again or explicitly promote the plan against the latest snapshot token.',
			},
		}
	}

	protected currentReadGenerations(): import('./types.ts').WorkbookGenerationInfo {
		return { workbook: 0, sheetMetadata: 0, formulas: 0, styles: 0 }
	}

	protected currentReadSnapshot(): WorkbookReadSnapshotInfo {
		const generations = this.currentReadGenerations()
		return {
			token: readSnapshotToken(generations, this.loadInfo),
			generations,
			load: this.loadInfo,
		}
	}

	inspect(): WorkbookInfo {
		if (this.workbookInfoCache) return this.workbookInfoCache
		let totalCells = 0
		let totalComments = 0
		let totalThreadedComments = 0
		let totalConditionalFormats = 0
		let totalDataValidations = 0
		let totalX14ConditionalFormats = 0
		let totalX14DataValidations = 0
		let totalImages = 0
		let totalSparklineGroups = 0
		let totalAdvancedFilters = 0
		const sheets = this.wb.sheets.map((sheet) => {
			const cellsHydrated = this.loadInfo.cellsHydrated
			const richSheetMetadataHydrated = this.loadInfo.richSheetMetadataHydrated
			const used = cellsHydrated ? sheet.cells.usedRange() : null
			const count = cellsHydrated ? sheet.cells.cellCount() : null
			if (count !== null) totalCells += count
			if (richSheetMetadataHydrated) {
				totalComments += sheet.comments.size
				totalThreadedComments += sheet.threadedComments.length
				totalConditionalFormats += sheet.conditionalFormats.length
				totalDataValidations += sheet.dataValidations.length
				totalX14ConditionalFormats += sheet.x14ConditionalFormats.length
				totalX14DataValidations += sheet.x14DataValidations.length
				totalImages += sheet.imageRefs.length
				totalSparklineGroups += sheet.sparklineGroups.length
				totalAdvancedFilters += sheet.advancedFilters.length
			}
			return buildSheetInfo(sheet, cellsHydrated, richSheetMetadataHydrated, used, count)
		})
		const info: WorkbookInfo = {
			sheetCount: this.loadInfo.sourceSheets.length,
			loadedSheetCount: this.loadInfo.loadedSheets.length,
			sheets,
			definedNames: this.wb.definedNames.workbookKeys(),
			definedNameDetails: this.definedNames(),
			cellCount: this.loadInfo.cellsHydrated ? totalCells : null,
			commentCount: this.loadInfo.richSheetMetadataHydrated ? totalComments : null,
			threadedCommentCount: this.loadInfo.richSheetMetadataHydrated ? totalThreadedComments : null,
			conditionalFormatCount: this.loadInfo.richSheetMetadataHydrated
				? totalConditionalFormats
				: null,
			dataValidationCount: this.loadInfo.richSheetMetadataHydrated ? totalDataValidations : null,
			x14ConditionalFormatCount: this.loadInfo.richSheetMetadataHydrated
				? totalX14ConditionalFormats
				: null,
			x14DataValidationCount: this.loadInfo.richSheetMetadataHydrated
				? totalX14DataValidations
				: null,
			imageCount: this.loadInfo.richSheetMetadataHydrated ? totalImages : null,
			sparklineGroupCount: this.loadInfo.richSheetMetadataHydrated ? totalSparklineGroups : null,
			advancedFilterCount: this.loadInfo.richSheetMetadataHydrated ? totalAdvancedFilters : null,
			chartCount: this.wb.chartParts.length,
			chartSheetCount: this.wb.chartSheets.length,
			macroSheetCount: this.wb.macroSheets.length,
			pivotTableCount: this.wb.pivotTables.length,
			pivotCacheCount: this.wb.pivotCaches.length,
			pivotRefreshPlans: buildPivotRefreshPlans(this.wb, this.loadInfo.cellsHydrated),
			refreshMetadata: buildWorkbookRefreshMetadata(
				this.wb,
				this.compat,
				this.loadInfo.cellsHydrated,
			),
			slicerCount: this.wb.slicers.length,
			slicerCacheCount: this.wb.slicerCaches.length,
			timelineCount: this.wb.timelines.length,
			timelineCacheCount: this.wb.timelineCaches.length,
			connectionPartCount: this.wb.connectionParts.length,
			dataModelPartCount: this.wb.dataModelParts.length,
			activeContentCount: this.wb.activeContent.length,
			sourceFormat: this.compat.sourceFormat,
			workbookViewCount: this.wb.workbookViews.length,
			externalReferenceCount: this.wb.externalReferences.length,
			workbookViews: this.wb.workbookViews.map((view) => ({ ...view })),
			externalReferences: [...this.wb.externalReferences],
			externalReferenceDetails: this.wb.externalReferenceDetails.map((entry) => ({
				...entry,
			})),
			externalReferenceUsages: buildExternalReferenceUsages(this.wb, this.formulaAnalysis()),
			charts: this.wb.chartParts.map(copyChartInfo),
			chartSheets: this.wb.chartSheets.map(copyChartSheetInfo),
			macroSheets: this.wb.macroSheets.map((entry) => ({ ...entry })),
			hasWorkbookProtection: this.wb.workbookProtection !== null,
			pivotTables: this.wb.pivotTables.map(copyPivotTableInfo),
			pivotCaches: this.wb.pivotCaches.map(copyPivotCacheInfo),
			slicerCaches: this.wb.slicerCaches.map(copySlicerCacheInfo),
			slicers: this.wb.slicers.map((entry) => ({ ...entry })),
			timelineCaches: this.wb.timelineCaches.map((entry) => ({
				...entry,
				pivotTableNames: [...entry.pivotTableNames],
				...(entry.state
					? {
							state: {
								...entry.state,
								...(entry.state.selection ? { selection: { ...entry.state.selection } } : {}),
								...(entry.state.bounds ? { bounds: { ...entry.state.bounds } } : {}),
							},
						}
					: {}),
			})),
			timelines: this.wb.timelines.map((entry) => ({ ...entry })),
			connectionParts: this.wb.connectionParts.map((entry) => ({ ...entry })),
			dataModelParts: this.wb.dataModelParts.map((entry) => ({ ...entry })),
			activeContent: this.wb.activeContent.map(cloneActiveContentInfo),
			documentProperties: copyDocumentProperties(this.wb.documentProperties),
			styleSummary: { ...this.wb.styleMetadata },
			themeSummary: {
				hasThemePart: this.wb.preservedTheme !== null,
				colors: this.wb.themeColors.map((color) => ({ ...color })),
				...this.wb.themeMetadata,
			},
			capabilityWarnings: buildCapabilityWarnings(this.wb, this.formulaAnalysis()),
			compatibility: this.compat,
			load: this.loadInfo,
		}
		this.workbookInfoCache = info
		return info
	}

	trustReport(options?: WorkbookTrustReportOptions): WorkbookTrustReport {
		return buildWorkbookTrustReport(this.wb, this.inspect(), options)
	}

	visualInventory(): WorkbookVisualInventoryInfo {
		let totalImages = 0
		let totalDrawingObjects = 0
		const charts = this.wb.chartParts.map(copyChartInfo)
		const chartSheets = this.wb.chartSheets.map(copyChartSheetInfo)
		const sheets = this.wb.sheets.map((sheet) => {
			const drawingRefs = this.loadInfo.cellsHydrated ? { ...sheet.drawingRefs } : null
			const imageRefs = this.loadInfo.richSheetMetadataHydrated
				? sheet.imageRefs.map(cloneSheetImageRef)
				: null
			const drawingObjectRefs = this.loadInfo.richSheetMetadataHydrated
				? sheet.drawingObjectRefs.map(cloneSheetDrawingObjectRef)
				: null
			if (imageRefs) totalImages += imageRefs.length
			if (drawingObjectRefs) totalDrawingObjects += drawingObjectRefs.length
			return {
				sheet: sheet.name,
				drawingRefs,
				hasDrawing: drawingRefs ? drawingRefs.hasDrawing : null,
				hasLegacyDrawing: drawingRefs ? drawingRefs.hasLegacyDrawing : null,
				imageRefs,
				drawingObjectRefs,
				imageCount: imageRefs ? imageRefs.length : null,
				drawingObjectCount: drawingObjectRefs ? drawingObjectRefs.length : null,
			}
		})
		const packageFeatures = this.compat.features.flatMap((feature) => {
			const category = classifyVisualFeature(feature.feature)
			return category ? [{ ...feature, category, locations: [...feature.locations] }] : []
		})
		const packageChartFeatureCount = packageFeatures
			.filter((feature) => feature.category === 'chart')
			.reduce((sum, feature) => sum + feature.count, 0)
		const packageChartSidecarFeatureCount = packageFeatures
			.filter(
				(feature) =>
					feature.feature === 'preservedChartStyle' || feature.feature === 'preservedChartColor',
			)
			.reduce((sum, feature) => sum + feature.count, 0)
		const packageDrawingFeatureCount = packageFeatures
			.filter(
				(feature) => feature.category === 'drawing' || feature.category === 'shape-or-control',
			)
			.reduce((sum, feature) => sum + feature.count, 0)
		const packageMediaFeatureCount = packageFeatures
			.filter((feature) => feature.category === 'image')
			.reduce((sum, feature) => sum + feature.count, 0)
		const notes: string[] = []
		if (!this.loadInfo.cellsHydrated) notes.push('Drawing references require full sheet hydration.')
		if (!this.loadInfo.richSheetMetadataHydrated)
			notes.push('Image and drawing-object references require rich sheet metadata hydration.')
		if (charts.length > 0)
			notes.push('Chart parts expose type, title, and series source refs; source edits are staged.')
		else if (packageChartFeatureCount > 0)
			notes.push('Chart parts are preserved but not structurally parsed in this load mode.')
		if (packageChartSidecarFeatureCount > 0)
			notes.push('Chart style/color sidecars are preserved separately from chart definitions.')
		if (chartSheets.length > 0)
			notes.push('Chartsheets are inventoried and blocked by the loss audit before writes.')
		if (packageDrawingFeatureCount > 0)
			notes.push(
				'Drawing and shape parts are currently preserve-first except parsed image anchors.',
			)
		return {
			load: this.loadInfo,
			packageFeatures,
			sheets,
			sheetImageCount: this.loadInfo.richSheetMetadataHydrated ? totalImages : null,
			sheetDrawingObjectCount: this.loadInfo.richSheetMetadataHydrated ? totalDrawingObjects : null,
			charts,
			chartSheets,
			structuredChartCount: charts.length,
			chartSheetCount: chartSheets.length,
			packageChartFeatureCount,
			packageChartSidecarFeatureCount,
			packageDrawingFeatureCount,
			packageMediaFeatureCount,
			hasPreservedCharts: packageFeatures.some(
				(feature) => feature.category === 'chart' && feature.tier === 'preserved',
			),
			hasPreservedChartSidecars: packageFeatures.some(
				(feature) =>
					(feature.feature === 'preservedChartStyle' ||
						feature.feature === 'preservedChartColor') &&
					feature.tier === 'preserved',
			),
			hasPreservedDrawings: packageFeatures.some(
				(feature) =>
					(feature.category === 'drawing' || feature.category === 'shape-or-control') &&
					feature.tier === 'preserved',
			),
			hasPreservedMedia: packageFeatures.some(
				(feature) => feature.category === 'image' && feature.tier === 'preserved',
			),
			notes,
		}
	}

	inspectSheet(name: string): SheetInspectInfo | undefined {
		if (this.sheetInspectCache.has(name)) return this.sheetInspectCache.get(name)
		const sheet = this.wb.getSheet(name)
		if (!sheet) {
			this.sheetInspectCache.set(name, undefined)
			return undefined
		}
		const cellsHydrated = this.loadInfo.cellsHydrated
		const richSheetMetadataHydrated = this.loadInfo.richSheetMetadataHydrated
		const used = cellsHydrated ? sheet.cells.usedRange() : null
		const count = cellsHydrated ? sheet.cells.cellCount() : null
		const base = buildSheetInfo(sheet, cellsHydrated, richSheetMetadataHydrated, used, count)
		const info = {
			...base,
			usedRange: used,
			state: sheet.state,
			merges: cellsHydrated ? [...sheet.merges] : null,
			tables: cellsHydrated ? sheet.tables.map((table) => buildTableInfo(table, sheet)) : null,
			comments: richSheetMetadataHydrated
				? [...sheet.comments.entries()].map(([ref, comment]) => ({ ref, ...comment }))
				: null,
			threadedComments: richSheetMetadataHydrated
				? sheet.threadedComments.map((comment) => ({ ...comment }))
				: null,
			hyperlinks: richSheetMetadataHydrated
				? [...sheet.hyperlinks.entries()].map(([ref, hyperlink]) => ({ ref, ...hyperlink }))
				: null,
			ignoredErrors: cellsHydrated ? [...sheet.ignoredErrors] : null,
			conditionalFormats: richSheetMetadataHydrated ? [...sheet.conditionalFormats] : null,
			dataValidations: richSheetMetadataHydrated ? [...sheet.dataValidations] : null,
			x14ConditionalFormats: richSheetMetadataHydrated
				? sheet.x14ConditionalFormats.map(cloneX14ConditionalFormatInfo)
				: null,
			x14DataValidations: richSheetMetadataHydrated
				? sheet.x14DataValidations.map(cloneX14DataValidationInfo)
				: null,
			imageRefs: richSheetMetadataHydrated ? sheet.imageRefs.map(cloneSheetImageRef) : null,
			drawingObjectRefs: richSheetMetadataHydrated
				? sheet.drawingObjectRefs.map(cloneSheetDrawingObjectRef)
				: null,
			sparklineGroups: richSheetMetadataHydrated
				? sheet.sparklineGroups.map((group) => ({
						...group,
						...(group.sparklines
							? { sparklines: group.sparklines.map((sparkline) => ({ ...sparkline })) }
							: {}),
					}))
				: null,
			advancedFilters: richSheetMetadataHydrated
				? sheet.advancedFilters.map((filter) => ({
						...filter,
						...(filter.autoFilter ? { autoFilter: copyAutoFilterInfo(filter.autoFilter) } : {}),
					}))
				: null,
			drawingRefs: cellsHydrated ? { ...sheet.drawingRefs } : null,
			autoFilter: cellsHydrated ? sheet.autoFilter : null,
			protection: cellsHydrated ? sheet.protection : null,
			protectedRanges: cellsHydrated ? sheet.protectedRanges.map((range) => ({ ...range })) : null,
			tabColor: cellsHydrated ? sheet.tabColor : null,
			sheetFormatPr: cellsHydrated ? sheet.sheetFormatPr : null,
			pageMargins: cellsHydrated ? sheet.pageMargins : null,
			pageSetup: cellsHydrated ? sheet.pageSetup : null,
			printOptions: cellsHydrated ? sheet.printOptions : null,
			headerFooter: cellsHydrated ? sheet.headerFooter : null,
		}
		this.sheetInspectCache.set(name, info)
		return info
	}

	sheet(name: string): SheetHandle | undefined {
		const cached = this.sheetHandleCache.get(name)
		if (cached) return cached
		const sheet = this.wb.getSheet(name)
		if (!sheet) return undefined
		const sheetIndex = this.wb.sheets.findIndex((s) => s.name === name)
		const handle = new SheetHandle(
			name,
			() => this.wb.getSheet(name),
			(row, col, cell) =>
				sheetIndex >= 0
					? resolveCellFormulaText(this.wb, sheetIndex, row, col, cell)
					: cell.formula,
			(_row, _col, cell, options) =>
				formatStyledDisplayCellValue(cell.value, this.wb.styles.get(cell.styleId), {
					dateSystem: this.wb.calcSettings.dateSystem,
					...options,
				}),
			() => this.currentReadSnapshot(),
		)
		this.sheetHandleCache.set(name, handle)
		return handle
	}

	readRange(sheetName: string, range: string): import('./types.ts').RangeInfo | undefined {
		return this.sheet(sheetName)?.range(range)
	}

	readRangeCompact(
		sheetName: string,
		range: string,
		opts?: { includeRefs?: boolean; omitEmpty?: boolean; flatValues?: boolean },
	): CompactRangeInfo | undefined {
		return this.sheet(sheetName)?.rangeCompact(range, opts)
	}

	cellStyle(cellRef: CellSelector): CellStyle | undefined {
		const { sheetName, ref } = normalizeCellSelector(cellRef, this.wb)
		const sheet = this.wb.getSheet(sheetName)
		if (!sheet) return undefined
		const parsed = parseA1(ref)
		const styleId = sheet.cells.readStyleId(parsed.row, parsed.col)
		if (styleId === undefined) return undefined
		const style = this.wb.styles.get(styleId)
		return style ? cloneStyle(style) : undefined
	}

	formatCellForDisplay(cellRef: CellSelector, options?: FormatDisplayOptions): string | undefined {
		const { sheetName, ref } = normalizeCellSelector(cellRef, this.wb)
		return this.sheet(sheetName)?.formatCellForDisplay(ref, options)
	}

	readWindow(
		sheetName: string,
		range: string,
		opts?: { rowOffset?: number; rowLimit?: number },
	): RangeWindowInfo | undefined {
		return this.sheet(sheetName)?.readWindow(range, opts)
	}

	readWindowCompact(
		sheetName: string,
		range: string,
		opts?: AgentReadOptions,
	): CompactRangeWindowInfo | undefined {
		return this.sheet(sheetName)?.readWindowCompact(range, opts)
	}

	readRows(
		sheetName: string,
		range: string,
		opts?: { rowOffset?: number; rowLimit?: number },
	): RangeRowsInfo | undefined {
		return this.sheet(sheetName)?.readRows(range, opts)
	}

	readObjects(
		sheetName: string,
		range: string,
		opts?: { rowOffset?: number; rowLimit?: number; headers?: readonly string[] | 'first-row' },
	): RangeObjectsInfo | undefined {
		return this.sheet(sheetName)?.readObjects(range, opts)
	}

	aggregateRange(
		sheetName: string,
		range: string,
	): import('./types.ts').RangeAggregateInfo | undefined {
		return this.sheet(sheetName)?.aggregateRange(range)
	}

	agentView(
		sheetName: string,
		range: string,
		opts?: AgentViewOptions,
	): AgentViewResult | undefined {
		const sheet = this.sheet(sheetName)
		if (!sheet) return undefined
		const rowChunkSize = Math.max(1, opts?.rowChunkSize ?? 256)
		const sampleRowLimit = Math.max(1, opts?.sampleRowLimit ?? 8)
		const sampleValueLimit = Math.max(1, opts?.sampleValueLimit ?? 4)
		let requestedRef: RangeRef | null = null
		let rowCount = 0
		let colCount = 0
		let nonEmptyCount = 0
		let formulaCount = 0
		const distinctFunctions = new Set<string>()
		const formulaPatterns = new Map<string, number>()
		const formulaPatternExamples = new Map<string, string[]>()
		const columnKinds = new Map<number, Set<AgentColumnSummary['kind']>>()
		const columnNonEmpty = new Map<number, number>()
		const columnFormulaCount = new Map<number, number>()
		const columnSamples = new Map<number, FlatCellValue[]>()
		const firstRowValues = new Map<number, FlatCellValue | null>()
		const samples: AgentSampleRow[] = []

		for (const window of this.streamWindowsCompact(sheetName, range, {
			rowLimit: rowChunkSize,
			includeRefs: true,
		})) {
			if (!requestedRef) {
				requestedRef = window.requestedRef
				rowCount = window.requestedRef.end.row - window.requestedRef.start.row + 1
				colCount = window.requestedRef.end.col - window.requestedRef.start.col + 1
			}
			const rows = groupCompactCellsByRow(window.cells)
			for (const [row, cells] of rows) {
				if (samples.length < sampleRowLimit && cells.length > 0) {
					samples.push({ row, cells })
				}
			}
			for (const cell of window.cells) {
				if (cell.value.kind !== 'empty') {
					nonEmptyCount++
					columnNonEmpty.set(cell.col, (columnNonEmpty.get(cell.col) ?? 0) + 1)
				}
				if (cell.formula !== null) {
					formulaCount++
					columnFormulaCount.set(cell.col, (columnFormulaCount.get(cell.col) ?? 0) + 1)
					for (const fn of extractFormulaFunctions(cell.formula)) distinctFunctions.add(fn)
					const pattern = normalizeFormulaPattern(cell.formula)
					formulaPatterns.set(pattern, (formulaPatterns.get(pattern) ?? 0) + 1)
					const examples = formulaPatternExamples.get(pattern)
					const ref = cell.ref ?? `${indexToColumn(cell.col)}${cell.row + 1}`
					if (examples) {
						if (examples.length < 3) examples.push(ref)
					} else {
						formulaPatternExamples.set(pattern, [ref])
					}
				}
				const kind: AgentColumnSummary['kind'] =
					cell.formula !== null
						? 'formula'
						: cell.value.kind === 'number' || cell.value.kind === 'date'
							? 'number'
							: cell.value.kind === 'string'
								? 'string'
								: cell.value.kind === 'boolean'
									? 'boolean'
									: cell.value.kind === 'empty'
										? 'empty'
										: 'mixed'
				let kinds = columnKinds.get(cell.col)
				if (!kinds) {
					kinds = new Set()
					columnKinds.set(cell.col, kinds)
				}
				kinds.add(kind)
				const flat = flattenForAgent(cell.value)
				if (flat !== null && (columnSamples.get(cell.col)?.length ?? 0) < sampleValueLimit) {
					const bucket = columnSamples.get(cell.col)
					if (bucket) bucket.push(flat)
					else columnSamples.set(cell.col, [flat])
				}
				if (requestedRef && cell.row === requestedRef.start.row && !firstRowValues.has(cell.col)) {
					firstRowValues.set(cell.col, flat)
				}
			}
		}

		if (!requestedRef) return undefined
		const columns: AgentColumnSummary[] = []
		for (let col = requestedRef.start.col; col <= requestedRef.end.col; col++) {
			const kinds = columnKinds.get(col) ?? new Set<AgentColumnSummary['kind']>(['empty'])
			let kind: AgentColumnSummary['kind'] = 'empty'
			const nonEmptyKinds = [...kinds].filter((entry) => entry !== 'empty')
			if ((columnFormulaCount.get(col) ?? 0) > 0 && (columnNonEmpty.get(col) ?? 0) === 0)
				kind = 'formula'
			else if (nonEmptyKinds.length === 1) kind = nonEmptyKinds[0] as AgentColumnSummary['kind']
			else if (nonEmptyKinds.length > 1) kind = 'mixed'
			columns.push({
				col,
				ref: indexToColumn(col),
				header: firstRowValues.get(col) ?? null,
				kind,
				nonEmptyCount: columnNonEmpty.get(col) ?? 0,
				formulaCount: columnFormulaCount.get(col) ?? 0,
				sampleValues: columnSamples.get(col) ?? [],
			})
		}
		const notes: string[] = []
		const density = rowCount * colCount > 0 ? nonEmptyCount / (rowCount * colCount) : 0
		if (density < 0.2) notes.push('Sparse window with many empty cells.')
		if (formulaCount > 0)
			notes.push(`${formulaCount} formula cells detected in the requested range.`)
		const topPattern = [...formulaPatterns.entries()].sort((a, b) => b[1] - a[1])[0]
		if (topPattern && topPattern[1] > 1) {
			notes.push(`Most common formula pattern repeats ${topPattern[1]} times.`)
		}

		const result: AgentViewResult = {
			sheet: sheetName,
			range: requestedRef,
			rowCount,
			colCount,
			nonEmptyCount,
			formulaCount,
			distinctFunctions: [...distinctFunctions].sort(),
			formulaPatterns: [...formulaPatterns.entries()]
				.sort((a, b) => b[1] - a[1])
				.slice(0, 12)
				.map(([pattern, count]) => ({
					pattern,
					count,
					examples: formulaPatternExamples.get(pattern) ?? [],
				})),
			columns,
			samples,
			notes,
		}
		return opts?.maxApproxTokens ? budgetAgentViewResult(result, opts.maxApproxTokens) : result
	}

	*streamRange(
		sheetName: string,
		range: string,
	): Generator<readonly import('./types.ts').CellInfo[]> {
		const sheet = this.sheet(sheetName)
		if (!sheet) return
		yield* sheet.streamRange(range)
	}

	*streamRangeCompact(
		sheetName: string,
		range: string,
		opts?: { includeRefs?: boolean },
	): Generator<readonly import('./types.ts').CompactCellInfo[]> {
		const sheet = this.sheet(sheetName)
		if (!sheet) return
		yield* sheet.streamRangeCompact(range, opts)
	}

	*streamWindows(
		sheetName: string,
		range: string,
		opts?: { rowLimit?: number },
	): Generator<RangeWindowInfo> {
		let rowOffset = 0
		while (true) {
			const window = this.readWindow(sheetName, range, {
				rowOffset,
				...(opts?.rowLimit !== undefined ? { rowLimit: opts.rowLimit } : {}),
			})
			if (!window) return
			yield window
			if (!window.hasMore || window.nextRowOffset === undefined) return
			rowOffset = window.nextRowOffset
		}
	}

	*streamWindowsCompact(
		sheetName: string,
		range: string,
		opts?: { rowLimit?: number; includeRefs?: boolean },
	): Generator<CompactRangeWindowInfo> {
		let rowOffset = 0
		while (true) {
			const window = this.readWindowCompact(sheetName, range, {
				rowOffset,
				...(opts?.rowLimit !== undefined ? { rowLimit: opts.rowLimit } : {}),
				...(opts?.includeRefs !== undefined ? { includeRefs: opts.includeRefs } : {}),
			})
			if (!window) return
			yield window
			if (!window.hasMore || window.nextRowOffset === undefined) return
			rowOffset = window.nextRowOffset
		}
	}

	table(name: string): TableHandle | undefined {
		const cached = this.tableHandleCache.get(name)
		if (cached) return cached
		for (const sheet of this.wb.sheets) {
			for (const table of sheet.tables) {
				if (table.name === name) {
					const handle = new TableHandle(name, () => {
						for (const currentSheet of this.wb.sheets) {
							const currentTable = currentSheet.tables.find((entry) => entry.name === name)
							if (currentTable) return { table: currentTable, sheet: currentSheet }
						}
						return undefined
					})
					this.tableHandleCache.set(name, handle)
					return handle
				}
			}
		}
		return undefined
	}

	pivotTables(sheetName?: string): readonly PivotTableInfo[] {
		return this.wb.pivotTables
			.filter((entry) => (sheetName ? entry.sheetName === sheetName : true))
			.map(copyPivotTableInfo)
	}

	pivotCaches(): readonly PivotCacheInfo[] {
		return this.wb.pivotCaches.map(copyPivotCacheInfo)
	}

	pivotCacheRows(options: PivotCacheRowsOptions = {}): readonly PivotCacheMaterializedRowInfo[] {
		return buildPivotCacheRows(this.wb.pivotCaches, options)
	}

	pivotOutputAudits(): readonly PivotOutputAuditInfo[] {
		return buildPivotOutputAudits(this.wb, this.loadInfo.cellsHydrated)
	}

	/**
	 * Build `setCells` operations that materialize supported PivotTable output cells.
	 * Unsupported layouts are reported in the result instead of guessed.
	 */
	pivotOutputMaterializeOps(
		options: PivotOutputMaterializeOptions = {},
	): PivotOutputMaterializeOpsResult {
		return buildPivotOutputMaterializeOps(this.wb, this.loadInfo.cellsHydrated, options)
	}

	pivotRefreshPlans(): readonly PivotRefreshPlanInfo[] {
		return buildPivotRefreshPlans(this.wb, this.loadInfo.cellsHydrated)
	}

	refreshMetadata(): WorkbookRefreshMetadataInfo {
		return buildWorkbookRefreshMetadata(this.wb, this.compat, this.loadInfo.cellsHydrated)
	}

	getPivotData(query: GetPivotDataQuery): GetPivotDataResult {
		return buildGetPivotDataResult(query, this.wb)
	}

	slicerCaches(): readonly SlicerCacheInfo[] {
		return this.wb.slicerCaches.map(copySlicerCacheInfo)
	}

	slicers(): readonly SlicerInfo[] {
		return this.wb.slicers.map((entry) => ({ ...entry }))
	}

	timelineCaches(): readonly TimelineCacheInfo[] {
		return this.wb.timelineCaches.map((entry) => ({
			...entry,
			pivotTableNames: [...entry.pivotTableNames],
			...(entry.state
				? {
						state: {
							...entry.state,
							...(entry.state.selection ? { selection: { ...entry.state.selection } } : {}),
							...(entry.state.bounds ? { bounds: { ...entry.state.bounds } } : {}),
						},
					}
				: {}),
		}))
	}

	timelines(): readonly TimelineInfo[] {
		return this.wb.timelines.map((entry) => ({ ...entry }))
	}

	connectionParts(): readonly WorkbookConnectionPartInfo[] {
		return this.wb.connectionParts.map((entry) => ({ ...entry }))
	}

	dataModelParts(): readonly WorkbookDataModelPartInfo[] {
		return this.wb.dataModelParts.map((entry) => ({ ...entry }))
	}

	activeContent(): readonly ActiveContentInfo[] {
		return this.wb.activeContent.map(cloneActiveContentInfo)
	}

	trace(cellRef: CellSelector, opts?: { maxDepth?: number }): TraceResult | undefined {
		if (this.dependencyVerificationIssue()) return undefined
		const { sheetName, ref } = normalizeCellSelector(cellRef, this.wb)
		const result = verifyTrace(this.wb, sheetName, ref, opts, {
			formulas: this.formulaAnalysis(),
			dependencies: this.dependencyAnalysis(),
		})
		if (!result.ok) return undefined
		return {
			ref: `${sheetName}!${ref}`,
			formula: result.value.formula,
			value: result.value.value,
			precedents: result.value.precedents.map((node) => ({
				ref: `${node.sheet}!${node.ref}`,
				formula: node.formula,
				value: node.value,
				depth: node.depth,
			})),
			dependents: result.value.dependents.map((node) => ({
				ref: `${node.sheet}!${node.ref}`,
				formula: node.formula,
				value: node.value,
				depth: node.depth,
			})),
			dependsOn: result.value.precedents.map((node) => `${node.sheet}!${node.ref}`),
			feedsInto: result.value.dependents.map((node) => `${node.sheet}!${node.ref}`),
		}
	}

	traceIssue(_cellRef: CellSelector): CheckIssue | undefined {
		const issue = this.dependencyVerificationIssue()
		return issue ? partialDependencyCheckIssue(issue) : undefined
	}

	formula(cellRef: CellSelector): FormulaInfo | undefined {
		const { sheetName, ref, cacheKey } = normalizeCellSelector(cellRef, this.wb)
		if (this.formulaInfoCache.has(cacheKey)) return this.formulaInfoCache.get(cacheKey)
		const sheet = this.wb.getSheet(sheetName)
		if (!sheet) {
			this.formulaInfoCache.set(cacheKey, undefined)
			return undefined
		}
		const parsedRef = parseA1(ref)
		const rawCell = sheet.cells.get(parsedRef.row, parsedRef.col)
		if (!rawCell) {
			this.formulaInfoCache.set(cacheKey, undefined)
			return undefined
		}
		const resolvedFormula = resolveCellFormulaText(
			this.wb,
			this.wb.sheets.findIndex((entry) => entry.id === sheet.id),
			parsedRef.row,
			parsedRef.col,
			rawCell,
		)
		if (!resolvedFormula) {
			this.formulaInfoCache.set(cacheKey, undefined)
			return undefined
		}

		const formula = normalizeFormulaInput(resolvedFormula)
		const tokens = tokenizeFormulaInput(formula)
		const parsed = parseFormula(formula)
		if (!parsed.ok) {
			const info = buildFormulaInfo({
				ref: `${sheetName}!${ref}`,
				formula,
				value: rawCell.value,
				binding: rawCell.formulaInfo ?? undefined,
				tokens,
				normalizedFormula: formula,
				functions: [],
				volatile: false,
				parseError: parsed.error.message,
			})
			this.formulaInfoCache.set(cacheKey, info)
			return info
		}

		const info = buildFormulaInfo({
			ref: `${sheetName}!${ref}`,
			formula,
			value: rawCell.value,
			binding: rawCell.formulaInfo ?? undefined,
			tokens,
			ast: parsed.value,
			normalizedFormula: printFormula(parsed.value),
			functions: [...collectFunctionNames(parsed.value)],
			volatile: hasVolatileFunction(parsed.value),
			references: collectFormulaReferences(parsed.value),
		})
		this.formulaInfoCache.set(cacheKey, info)
		return info
	}

	diff(other: WorkbookReadView): WorkbookDiff {
		return diffWorkbooks(this.wb, other.wb)
	}

	snapshot(): WorkbookSnapshot {
		return createSnapshot(this.wb)
	}

	toJSON(): object {
		const snap = createSnapshot(this.wb)
		return {
			sheets: snap.sheets,
			names: snap.names,
			calcSettings: this.wb.calcSettings,
			report: this.compat,
		}
	}

	get report(): CompatibilityReport {
		return this.compat
	}

	get sheets(): readonly string[] {
		return this.wb.sheets.map((sheet) => sheet.name)
	}

	get names(): readonly string[] {
		return this.wb.definedNames.workbookKeys()
	}

	dependencyVerificationIssue(): string | undefined {
		const reasons: string[] = []
		if (!this.loadInfo.hasAllSheets) reasons.push('not all sheets are loaded')
		if (this.loadInfo.maxRows !== undefined) {
			reasons.push(`only the first ${this.loadInfo.maxRows} row(s) are hydrated per loaded sheet`)
		}
		if (!this.loadInfo.cellsHydrated) reasons.push('sheet cells are not hydrated')
		if (this.loadInfo.mode === 'values') reasons.push('only cell values are hydrated')
		if (reasons.length === 0) return undefined
		return `Cannot verify workbook dependencies from this partial view because ${reasons.join(' and ')}. Reopen with formula or full mode and all sheets loaded.`
	}

	definedNames(scopeSheetName?: string): readonly DefinedNameInfo[] {
		if (scopeSheetName) {
			return this.wb.definedNames
				.list()
				.filter((entry) => {
					if (entry.scope.kind !== 'sheet') return false
					const scope = entry.scope
					const sheet = this.wb.sheets.find((candidate) => candidate.id === scope.sheetId)
					return sheet?.name === scopeSheetName
				})
				.map((entry) => buildDefinedNameInfo(this.wb, entry))
		}
		return this.wb.definedNames.list().map((entry) => buildDefinedNameInfo(this.wb, entry))
	}

	workbookViews(): readonly import('./types.ts').WorkbookViewInfo[] {
		return this.wb.workbookViews.map((view) => ({ ...view }))
	}

	externalReferences(): readonly string[] {
		return [...this.wb.externalReferences]
	}

	externalReferenceUsages(): readonly ExternalReferenceUsageInfo[] {
		return buildExternalReferenceUsages(this.wb, this.formulaAnalysis())
	}

	definedName(name: string, scopeSheetName?: string): DefinedNameInfo | undefined {
		let entry = scopeSheetName
			? resolveDefinedNameBySheet(this.wb, name, scopeSheetName)
			: this.wb.definedNames.getEntry(name)

		if (!entry && !scopeSheetName) {
			entry = this.wb.definedNames.list().find((definedName) => definedName.name === name)
		}
		if (!entry) return undefined

		const sheetScope = entry.scope.kind === 'sheet' ? entry.scope : undefined
		const sheetName = sheetScope
			? this.wb.sheets.find((sheet) => sheet.id === sheetScope.sheetId)?.name
			: undefined
		return {
			...buildDefinedNameInfo(this.wb, entry),
			scope: entry.scope.kind,
			...(sheetName ? { sheet: sheetName } : {}),
		}
	}

	protected clearReadCaches(): void {
		this.workbookInfoCache = undefined
		this.sheetInspectCache.clear()
		this.formulaInfoCache.clear()
		this.sheetHandleCache.clear()
		this.tableHandleCache.clear()
	}

	analysis(): WorkbookAnalysis {
		return analyzeWorkbook(this.wb)
	}

	dependencyAnalysis(): WorkbookDependencyAnalysis {
		return analyzeWorkbookDependencies(this.wb)
	}

	formulaAnalysis(): WorkbookFormulaAnalysis {
		return analyzeWorkbookFormulas(this.wb)
	}

	replaceWorkbook(
		workbook: Workbook,
		report: CompatibilityReport = this.compat,
		loadInfo: WorkbookLoadInfo = this.loadInfo,
	): void {
		this.wb = workbook
		this.compat = report
		this.loadInfo = loadInfo
		this.clearReadCaches()
	}
}

function buildCapabilityWarnings(
	workbook: Workbook,
	analysis: WorkbookFormulaAnalysis,
): CapabilityWarningInfo[] {
	const warnings: CapabilityWarningInfo[] = []
	const add = (capabilityId: string, evidence: readonly string[]): void => {
		if (evidence.length === 0) return
		const capability = getCapability(capabilityId)
		if (!capability || !isCapabilityGap(capability.status)) return
		warnings.push({
			capabilityId: capability.id,
			label: capability.label,
			family: capability.family,
			status: capability.status,
			priority: capability.priority,
			reason: capability.gapReason,
			nextMilestone: capability.nextMilestone,
			evidence,
		})
	}

	add('workbook.xlsm-package', [
		...(workbook.macroSheets.length > 0 ? [`macroSheets=${workbook.macroSheets.length}`] : []),
		...countActiveKinds(workbook.activeContent, ['vbaProject', 'macroSheet']).map(
			([kind, count]) => `${kind}=${count}`,
		),
	])
	add('analytics.pivots', [
		...(workbook.pivotTables.length > 0 ? [`pivotTables=${workbook.pivotTables.length}`] : []),
	])
	add('analytics.data-model', [
		...(workbook.dataModelParts.length > 0
			? [`dataModelParts=${workbook.dataModelParts.length}`]
			: []),
	])
	add('analytics.getpivotdata', getPivotDataEvidence(analysis))
	add(
		'active.vba-macros',
		countActiveKinds(workbook.activeContent, ['vbaProject']).map(
			([kind, count]) => `${kind}=${count}`,
		),
	)
	add(
		'active.activex-controls',
		countActiveKinds(workbook.activeContent, ['activeX']).map(
			([kind, count]) => `${kind}=${count}`,
		),
	)
	add(
		'active.form-controls',
		countActiveKinds(workbook.activeContent, ['formControl']).map(
			([kind, count]) => `${kind}=${count}`,
		),
	)
	add(
		'active.signatures',
		countActiveKinds(workbook.activeContent, ['vbaSignature', 'digitalSignature']).map(
			([kind, count]) => `${kind}=${count}`,
		),
	)
	add(
		'active.custom-ui',
		countActiveKinds(workbook.activeContent, ['customUi']).map(
			([kind, count]) => `${kind}=${count}`,
		),
	)
	add(
		'active.shape-macros',
		countActiveKinds(workbook.activeContent, ['shapeMacro']).map(
			([kind, count]) => `${kind}=${count}`,
		),
	)
	add('connections.power-query', [
		...countConnectionKinds(workbook.connectionParts, ['powerQueryMashup']).map(
			([kind, count]) => `${kind}=${count}`,
		),
	])
	add('visuals.chartsheets', [
		...(workbook.chartSheets.length > 0 ? [`chartSheets=${workbook.chartSheets.length}`] : []),
	])
	return warnings
}

function countActiveKinds(
	entries: readonly ActiveContentInfo[],
	kinds: readonly ActiveContentInfo['kind'][],
): Array<readonly [ActiveContentInfo['kind'], number]> {
	const counts = new Map<ActiveContentInfo['kind'], number>()
	for (const entry of entries) {
		if (!kinds.includes(entry.kind)) continue
		counts.set(entry.kind, (counts.get(entry.kind) ?? 0) + 1)
	}
	return [...counts.entries()]
}

function countConnectionKinds(
	entries: readonly WorkbookConnectionPartInfo[],
	kinds: readonly WorkbookConnectionPartInfo['kind'][],
): Array<readonly [WorkbookConnectionPartInfo['kind'], number]> {
	const counts = new Map<WorkbookConnectionPartInfo['kind'], number>()
	for (const entry of entries) {
		if (!kinds.includes(entry.kind)) continue
		counts.set(entry.kind, (counts.get(entry.kind) ?? 0) + 1)
	}
	return [...counts.entries()]
}

function getPivotDataEvidence(analysis: WorkbookFormulaAnalysis): string[] {
	let count = 0
	for (const formula of analysis.formulas.values()) {
		if (!formula.ast) continue
		if (collectFunctionNames(formula.ast).has('GETPIVOTDATA')) count++
	}
	return count > 0 ? [`GETPIVOTDATA formulas=${count}`] : []
}

function classifyVisualFeature(
	feature: string,
): 'chart' | 'drawing' | 'image' | 'shape-or-control' | null {
	const normalized = feature.toLowerCase()
	if (normalized.includes('chart')) return 'chart'
	if (normalized.includes('media') || normalized.includes('image')) return 'image'
	if (normalized.includes('activex') || normalized.includes('control')) return 'shape-or-control'
	if (normalized.includes('drawing') || normalized.includes('vml') || normalized.includes('shape'))
		return 'drawing'
	return null
}

function readSnapshotToken(
	generations: import('./types.ts').WorkbookGenerationInfo,
	load: WorkbookLoadInfo,
): string {
	const loadFlags = [
		load.mode,
		load.isPartial ? 'partial' : 'full',
		load.cellsHydrated ? 'cells' : 'no-cells',
		load.richSheetMetadataHydrated ? 'rich' : 'lean',
		load.hasAllSheets ? 'all-sheets' : 'selected-sheets',
		`rows:${load.maxRows ?? 'all'}`,
		`loaded:${load.loadedSheets.join(',')}`,
	].join('|')
	return [
		`w${generations.workbook}`,
		`m${generations.sheetMetadata}`,
		`f${generations.formulas}`,
		`s${generations.styles}`,
		loadFlags,
	].join(':')
}

function copyDocumentProperties(
	properties: WorkbookDocumentProperties,
): WorkbookDocumentProperties {
	return {
		...(properties.core ? { core: { ...properties.core } } : {}),
		...(properties.app
			? {
					app: Object.fromEntries(
						Object.entries(properties.app).map(([key, value]) => [
							key,
							Array.isArray(value) ? [...value] : value,
						]),
					),
				}
			: {}),
		...(properties.custom
			? { custom: properties.custom.map((property) => ({ ...property })) }
			: {}),
	}
}

function buildSheetInfo(
	sheet: import('@ascend/core').Sheet,
	cellsHydrated: boolean,
	richSheetMetadataHydrated: boolean,
	used: RangeRef | null,
	count: number | null,
): SheetInfo {
	return {
		name: sheet.name,
		rowCount: used ? used.end.row + 1 : null,
		colCount: used ? used.end.col + 1 : null,
		cellCount: count,
		tableCount: cellsHydrated ? sheet.tables.length : null,
		commentCount: richSheetMetadataHydrated ? sheet.comments.size : null,
		threadedCommentCount: richSheetMetadataHydrated ? sheet.threadedComments.length : null,
		conditionalFormatCount: richSheetMetadataHydrated ? sheet.conditionalFormats.length : null,
		dataValidationCount: richSheetMetadataHydrated ? sheet.dataValidations.length : null,
		x14ConditionalFormatCount: richSheetMetadataHydrated
			? sheet.x14ConditionalFormats.length
			: null,
		x14DataValidationCount: richSheetMetadataHydrated ? sheet.x14DataValidations.length : null,
		hasFrozenPanes: cellsHydrated ? sheet.frozenRows > 0 || sheet.frozenCols > 0 : null,
		colWidthCount: cellsHydrated ? sheet.colWidths.size : null,
		imageCount: richSheetMetadataHydrated ? sheet.imageRefs.length : null,
		sparklineGroupCount: richSheetMetadataHydrated ? sheet.sparklineGroups.length : null,
		advancedFilterCount: richSheetMetadataHydrated ? sheet.advancedFilters.length : null,
		rowHeightCount: cellsHydrated ? sheet.rowHeights.size : null,
		hyperlinkCount: richSheetMetadataHydrated ? sheet.hyperlinks.size : null,
		ignoredErrorCount: cellsHydrated ? sheet.ignoredErrors.length : null,
		hasAutoFilter: cellsHydrated ? sheet.autoFilter !== null : null,
		hasDrawingRefs: cellsHydrated
			? sheet.drawingRefs.hasDrawing || sheet.drawingRefs.hasLegacyDrawing
			: null,
		hasPageMetadata: cellsHydrated
			? sheet.pageMargins !== null ||
				sheet.pageSetup !== null ||
				sheet.printOptions !== null ||
				sheet.headerFooter !== null
			: null,
		hasProtection: cellsHydrated ? sheet.protection !== null : null,
		cellDataLoaded: cellsHydrated,
	}
}

function copyAutoFilterInfo(filter: AutoFilter): AutoFilter {
	return {
		...filter,
		columns: filter.columns.map((column) => ({
			...column,
			...(column.values ? { values: [...column.values] } : {}),
			...(column.dateGroupItems
				? { dateGroupItems: column.dateGroupItems.map((item) => ({ ...item })) }
				: {}),
			...(column.customFilters
				? { customFilters: column.customFilters.map((entry) => ({ ...entry })) }
				: {}),
		})),
		...(filter.sortState
			? {
					sortState: {
						...filter.sortState,
						conditions: filter.sortState.conditions.map((condition) => ({ ...condition })),
					},
				}
			: {}),
	}
}

function buildTableInfo(
	table: import('@ascend/core').Table,
	sheet: import('@ascend/core').Sheet,
): TableInfo {
	const headerOffset = table.hasHeaders ? 1 : 0
	const totalOffset = table.hasTotals ? 1 : 0
	const headerRow = table.hasHeaders
		? Array.from({ length: table.ref.end.col - table.ref.start.col + 1 }, (_, offset) => {
				return sheet.cells.get(table.ref.start.row, table.ref.start.col + offset)?.value ?? EMPTY
			})
		: undefined
	const totalsRow = table.hasTotals
		? Array.from({ length: table.ref.end.col - table.ref.start.col + 1 }, (_, offset) => {
				return sheet.cells.get(table.ref.end.row, table.ref.start.col + offset)?.value ?? EMPTY
			})
		: undefined
	return {
		name: table.name,
		ref: table.ref,
		...(table.partPath ? { partPath: table.partPath } : {}),
		...(table.contentType ? { contentType: table.contentType } : {}),
		...(table.contentTypeSource ? { contentTypeSource: table.contentTypeSource } : {}),
		...(table.sourcePartPath ? { sourcePartPath: table.sourcePartPath } : {}),
		...(table.sourceRelationshipPart
			? { sourceRelationshipPart: table.sourceRelationshipPart }
			: {}),
		...(table.sourceRelationshipId ? { sourceRelationshipId: table.sourceRelationshipId } : {}),
		...(table.sourceRelationshipType
			? { sourceRelationshipType: table.sourceRelationshipType }
			: {}),
		...(table.sourceRelationshipRawType
			? { sourceRelationshipRawType: table.sourceRelationshipRawType }
			: {}),
		...(table.sourceRelationshipRawTarget
			? { sourceRelationshipRawTarget: table.sourceRelationshipRawTarget }
			: {}),
		...(table.sourceRelationshipResolvedTarget
			? { sourceRelationshipResolvedTarget: table.sourceRelationshipResolvedTarget }
			: {}),
		...(table.sourceRelationshipTargetMode
			? { sourceRelationshipTargetMode: table.sourceRelationshipTargetMode }
			: {}),
		rowCount: table.ref.end.row - table.ref.start.row + 1 - headerOffset - totalOffset,
		hasHeaders: table.hasHeaders,
		hasTotals: table.hasTotals,
		autoFilter: table.autoFilter ?? null,
		...(table.sortState ? { sortState: table.sortState } : {}),
		...(table.tableStyleInfo ? { styleInfo: table.tableStyleInfo } : {}),
		...(table.queryTable ? { queryTable: { ...table.queryTable } } : {}),
		columnDefs: [...table.columns],
		...(headerRow ? { headerRow } : {}),
		...(totalsRow ? { totalsRow } : {}),
	}
}

function buildDefinedNameInfo(
	workbook: Workbook,
	entry: ReturnType<Workbook['definedNames']['list']>[number],
): DefinedNameInfo {
	const formula = normalizeFormulaInput(entry.formula)
	const parsed = parseFormula(formula)
	if (!parsed.ok) {
		return {
			name: entry.name,
			formula: entry.formula,
			normalizedFormula: formula,
			scope: entry.scope.kind,
			references: [],
			refs: [],
			functions: [],
			volatile: false,
			parseError: parsed.error.message,
		}
	}
	const references = collectFormulaReferences(parsed.value)
	const sheetScope = entry.scope.kind === 'sheet' ? entry.scope : undefined
	const scopeSheet = sheetScope
		? workbook.sheets.find((sheet) => sheet.id === sheetScope.sheetId)?.name
		: undefined
	return {
		name: entry.name,
		formula: entry.formula,
		normalizedFormula: printFormula(parsed.value),
		scope: entry.scope.kind,
		...(scopeSheet ? { sheet: scopeSheet } : {}),
		references,
		refs: flattenLegacyReferenceTexts(references),
		functions: [...collectFunctionNames(parsed.value)],
		volatile: hasVolatileFunction(parsed.value),
	}
}

function resolveDefinedNameBySheet(
	workbook: Workbook,
	name: string,
	sheetName: string,
): ReturnType<Workbook['definedNames']['resolve']> {
	const sheet = workbook.getSheet(sheetName)
	if (!sheet) return undefined
	return (
		workbook.definedNames.resolve(name, sheet.id, sheet.id) ??
		workbook.definedNames.resolve(name, sheet.id)
	)
}

function copyChartInfo(chart: ChartPartInfo): ChartPartInfo {
	return {
		...chart,
		series: chart.series.map((series) => ({ ...series })),
	}
}

function copyChartSheetInfo(
	chartSheet: import('@ascend/core').ChartSheetInfo,
): import('@ascend/core').ChartSheetInfo {
	return {
		...chartSheet,
		chartPartPaths: [...chartSheet.chartPartPaths],
	}
}

function cloneSheetImageRef(ref: SheetImageRef): SheetImageRef {
	return {
		...ref,
		...(ref.content ? { content: new Uint8Array(ref.content) } : {}),
		...(ref.anchor ? { anchor: cloneSheetImageAnchor(ref.anchor) } : {}),
	}
}

function cloneSheetDrawingObjectRef(ref: SheetDrawingObjectRef): SheetDrawingObjectRef {
	return {
		...ref,
		...(ref.anchor ? { anchor: cloneSheetImageAnchor(ref.anchor) } : {}),
		...(ref.relIds ? { relIds: [...ref.relIds] } : {}),
		...(ref.relationshipRefs
			? { relationshipRefs: ref.relationshipRefs.map((relationship) => ({ ...relationship })) }
			: {}),
	}
}

function cloneSheetImageAnchor(anchor: SheetImageAnchor): SheetImageAnchor {
	switch (anchor.kind) {
		case 'oneCell':
			return { ...anchor, from: { ...anchor.from } }
		case 'twoCell':
			return { ...anchor, from: { ...anchor.from }, to: { ...anchor.to } }
		case 'absolute':
			return { ...anchor }
	}
}

function copyPivotCacheInfo(cache: PivotCacheInfo): PivotCacheInfo {
	return clonePivotCacheInfo(cache)
}

function buildPivotCacheRows(
	caches: readonly PivotCacheInfo[],
	options: PivotCacheRowsOptions,
): PivotCacheMaterializedRowInfo[] {
	const limit = options.limit === undefined ? Number.POSITIVE_INFINITY : Math.max(0, options.limit)
	const rows: PivotCacheMaterializedRowInfo[] = []
	for (const cache of caches) {
		if (options.cacheId !== undefined && cache.cacheId !== options.cacheId) continue
		if (options.partPath !== undefined && cache.partPath !== options.partPath) continue
		for (const record of cache.records?.materializedRecords ?? []) {
			if (rows.length >= limit) return rows
			rows.push({
				partPath: cache.partPath,
				...(cache.cacheId !== undefined ? { cacheId: cache.cacheId } : {}),
				rowIndex: record.index,
				values: record.values.map((value) => decodePivotCacheValue(cache, value)),
			})
		}
	}
	return rows
}

function decodePivotCacheValue(
	cache: PivotCacheInfo,
	value: PivotCacheRecordValueInfo,
): PivotCacheDecodedValueInfo {
	const field = cache.fields[value.index]
	const sharedItem =
		value.kind === 'sharedItem'
			? field?.sharedItems?.find((item) => item.index === value.sharedItemIndex)
			: undefined
	const decodedValue = value.value ?? sharedItem?.value
	return {
		fieldIndex: value.index,
		...(field?.name !== undefined ? { fieldName: field.name } : {}),
		rawKind: value.kind,
		kind: sharedItem?.kind ?? value.kind,
		...(decodedValue !== undefined ? { value: decodedValue } : {}),
		...(value.sharedItemIndex !== undefined ? { sharedItemIndex: value.sharedItemIndex } : {}),
		...(sharedItem?.kind !== undefined ? { sharedItemKind: sharedItem.kind } : {}),
	}
}

function buildPivotOutputAudits(
	workbook: Workbook,
	cellsHydrated: boolean,
): PivotOutputAuditInfo[] {
	const audits: PivotOutputAuditInfo[] = []
	const context = createPivotOutputAuditContext()
	for (const pivot of workbook.pivotTables) {
		const audit = buildPivotOutputAudit(workbook, cellsHydrated, pivot, context)
		if (audit) audits.push(audit)
	}
	return audits
}

function buildPivotOutputMaterializeOps(
	workbook: Workbook,
	cellsHydrated: boolean,
	options: PivotOutputMaterializeOptions,
): PivotOutputMaterializeOpsResult {
	const context = createPivotOutputAuditContext()
	const mode = options.mode ?? 'mismatches'
	const updatesBySheet = new Map<string, CellUpdate[]>()
	const unsupported: PivotOutputMaterializeUnsupportedInfo[] = []
	let plannedCellCount = 0
	for (const pivot of workbook.pivotTables) {
		if (options.pivotTable !== undefined && pivot.name !== options.pivotTable) continue
		if (options.partPath !== undefined && pivot.partPath !== options.partPath) continue
		const plan = buildPivotOutputMaterializePlan(workbook, cellsHydrated, pivot, context, mode)
		if (!plan.ok) {
			unsupported.push(plan.unsupported)
			continue
		}
		if (plan.updates.length === 0) continue
		plannedCellCount += plan.updates.length
		const sheetUpdates = updatesBySheet.get(plan.sheetName)
		if (sheetUpdates) {
			sheetUpdates.push(...plan.updates)
		} else {
			updatesBySheet.set(plan.sheetName, [...plan.updates])
		}
	}
	const ops: Operation[] = Array.from(updatesBySheet, ([sheet, updates]) => ({
		op: 'setCells',
		sheet,
		updates,
	}))
	return { ops, plannedCellCount, unsupported }
}

function buildPivotOutputMaterializePlan(
	workbook: Workbook,
	cellsHydrated: boolean,
	pivot: PivotTableInfo,
	context: PivotOutputAuditContext,
	mode: PivotOutputMaterializeMode,
):
	| { ok: true; sheetName: string; updates: readonly CellUpdate[] }
	| { ok: false; unsupported: PivotOutputMaterializeUnsupportedInfo } {
	const base = pivotMaterializeUnsupportedBase(pivot)
	if (!cellsHydrated) {
		return unsupportedPivotMaterialize(
			base,
			'Pivot output cells are not hydrated in this load mode.',
		)
	}
	const sheet = workbook.getSheet(pivot.sheetName)
	if (!sheet) return unsupportedPivotMaterialize(base, 'Pivot output sheet was not loaded.')
	if (!pivot.locationRef) {
		return unsupportedPivotMaterialize(base, 'Pivot table has no output location.')
	}
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return unsupportedPivotMaterialize(base, `Pivot output range is invalid: ${pivot.locationRef}`)
	}
	let cache = workbook.pivotCaches.find((entry) => entry.cacheId === pivot.cacheId)
	if (!cache) return unsupportedPivotMaterialize(base, 'Pivot cache metadata was not found.')
	if (isEmptyPivotOutput(pivot)) {
		return unsupportedPivotMaterialize(
			base,
			'Empty PivotTable outputs do not need materialization.',
		)
	}
	if (pivot.dataFields.length === 0) {
		return unsupportedPivotMaterialize(base, 'No-data PivotTable outputs are not materialized.')
	}
	if (pivot.rowFields.length !== 1) {
		return unsupportedPivotMaterialize(
			base,
			'Only one-row-field PivotTable outputs are materialized.',
		)
	}
	if (pivot.columnFields.length > 0 && !pivot.columnFields.every((field) => field.index === -2)) {
		return unsupportedPivotMaterialize(
			base,
			'Column-field PivotTable outputs are not materialized yet.',
		)
	}
	const rowFieldIndex = pivot.rowFields[0]?.index
	if (rowFieldIndex === undefined || rowFieldIndex < 0) {
		return unsupportedPivotMaterialize(base, 'Pivot row field metadata was not found.')
	}
	const auditCache = pivotCacheWithMaterializedAuditRows(workbook, cache)
	if (!auditCache.ok) return unsupportedPivotMaterialize(base, auditCache.warning)
	cache = auditCache.value
	const expected = aggregateSimplePivotOutput(
		cache,
		pivot,
		rowFieldIndex,
		pivotOutputAuditCacheRows(context, cache),
	)
	if (!expected.ok) return unsupportedPivotMaterialize(base, expected.warning)
	const updates = buildSimplePivotOutputMaterializeUpdates(
		sheet,
		bounds,
		pivot,
		expected.value,
		mode,
	)
	if (!updates.ok) return unsupportedPivotMaterialize(base, updates.warning)
	return { ok: true, sheetName: pivot.sheetName, updates: updates.value }
}

function pivotMaterializeUnsupportedBase(
	pivot: PivotTableInfo,
): Omit<PivotOutputMaterializeUnsupportedInfo, 'warning'> {
	return {
		...(pivot.name !== undefined ? { pivotTable: pivot.name } : {}),
		partPath: pivot.partPath,
		sheetName: pivot.sheetName,
		...(pivot.cacheId !== undefined ? { cacheId: pivot.cacheId } : {}),
	}
}

function unsupportedPivotMaterialize(
	base: Omit<PivotOutputMaterializeUnsupportedInfo, 'warning'>,
	warning: string,
): { ok: false; unsupported: PivotOutputMaterializeUnsupportedInfo } {
	return { ok: false, unsupported: { ...base, warning } }
}

type PivotOutputAuditContext = {
	readonly cacheRowsByKey: Map<string, readonly PivotCacheMaterializedRowInfo[]>
}

function createPivotOutputAuditContext(): PivotOutputAuditContext {
	return { cacheRowsByKey: new Map() }
}

function pivotOutputAuditCacheRows(
	context: PivotOutputAuditContext,
	cache: PivotCacheInfo,
): readonly PivotCacheMaterializedRowInfo[] {
	const key = pivotOutputAuditCacheKey(cache)
	const cached = context.cacheRowsByKey.get(key)
	if (cached) return cached
	const rows = buildPivotCacheRows([cache], {})
	context.cacheRowsByKey.set(key, rows)
	return rows
}

function pivotOutputAuditCacheKey(cache: PivotCacheInfo): string {
	return `${cache.partPath}\u0000${cache.cacheId ?? ''}`
}

function buildPivotOutputAudit(
	workbook: Workbook,
	cellsHydrated: boolean,
	pivot: PivotTableInfo,
	context: PivotOutputAuditContext,
): PivotOutputAuditInfo | undefined {
	const base = {
		...(pivot.name !== undefined ? { pivotTable: pivot.name } : {}),
		partPath: pivot.partPath,
		sheetName: pivot.sheetName,
		...(pivot.cacheId !== undefined ? { cacheId: pivot.cacheId } : {}),
	}
	if (!cellsHydrated) {
		return unsupportedPivotAudit(base, 'Pivot output cells are not hydrated in this load mode.')
	}
	const sheet = workbook.getSheet(pivot.sheetName)
	if (!sheet) return unsupportedPivotAudit(base, 'Pivot output sheet was not loaded.')
	if (!pivot.locationRef) return unsupportedPivotAudit(base, 'Pivot table has no output location.')
	let cache = workbook.pivotCaches.find((entry) => entry.cacheId === pivot.cacheId)
	if (!cache) return unsupportedPivotAudit(base, 'Pivot cache metadata was not found.')
	if (isEmptyPivotOutput(pivot)) {
		const emptyAudit = auditEmptyPivotOutput(workbook, sheet.id, pivot)
		if (!emptyAudit.ok) return unsupportedPivotAudit(base, emptyAudit.warning)
		return {
			...base,
			status: 'passed',
			checkedValueCount: emptyAudit.checkedCellCount,
			mismatches: [],
			warnings: [],
		}
	}
	if (!hasSavedSheetCells(sheet)) {
		return unsupportedPivotAudit(base, 'Pivot output sheet has no saved cells to audit.')
	}
	const auditCache = pivotCacheWithMaterializedAuditRows(workbook, cache)
	if (!auditCache.ok) return unsupportedPivotAudit(base, auditCache.warning)
	cache = auditCache.value
	const cacheRows = pivotOutputAuditCacheRows(context, cache)
	if (pivot.dataFields.length === 0) {
		const labelAudit = auditRowLabelOnlyPivotOutput(workbook, sheet.id, pivot, cache)
		if (!labelAudit.ok) return unsupportedPivotAudit(base, labelAudit.warning)
		return {
			...base,
			status: 'passed',
			checkedValueCount: labelAudit.checkedCellCount,
			mismatches: [],
			warnings: [],
		}
	}
	if (isDataFieldsNestedOnRowsPivot(pivot)) {
		const expected = aggregateDataFieldsNestedOnRowsPivotOutput(cache, pivot, cacheRows)
		if (!expected.ok) return unsupportedPivotAudit(base, expected.warning)
		const actual = readMultiRowAxisPivotOutput(
			workbook,
			sheet.id,
			pivot,
			expected.value.rowKeys,
			expected.value.columnKeys,
		)
		if (!actual.ok) return unsupportedPivotAudit(base, actual.warning)
		const mismatches = comparePivotOutput(expected.value.values, actual.value)
		return {
			...base,
			status: mismatches.length > 0 ? 'mismatch' : 'passed',
			checkedValueCount: expected.value.checkedValueCount,
			mismatches,
			warnings: [],
		}
	}
	if (isDataFieldsOnRowsPivot(pivot)) {
		const expected = aggregateDataFieldsOnRowsPivotOutput(cache, pivot, cacheRows)
		if (!expected.ok) return unsupportedPivotAudit(base, expected.warning)
		const actual = readDataFieldsOnRowsPivotOutput(
			workbook,
			sheet.id,
			pivot,
			expected.value.dataFieldNames,
			expected.value.columnKeys,
		)
		if (!actual.ok) return unsupportedPivotAudit(base, actual.warning)
		const mismatches = comparePivotOutput(expected.value.values, actual.value)
		return {
			...base,
			status: mismatches.length > 0 ? 'mismatch' : 'passed',
			checkedValueCount: expected.value.checkedValueCount,
			mismatches,
			warnings: [],
		}
	}
	if (isDataFieldsOnColumnsPivot(pivot)) {
		const expected = aggregateDataFieldsOnColumnsPivotOutput(cache, pivot, cacheRows)
		if (!expected.ok) return unsupportedPivotAudit(base, expected.warning)
		const actual = readMultiRowAxisPivotOutput(
			workbook,
			sheet.id,
			pivot,
			expected.value.rowKeys,
			expected.value.columnKeys,
		)
		if (!actual.ok) return unsupportedPivotAudit(base, actual.warning)
		const mismatches = comparePivotOutput(expected.value.values, actual.value)
		return {
			...base,
			status: mismatches.length > 0 ? 'mismatch' : 'passed',
			checkedValueCount: expected.value.checkedValueCount,
			mismatches,
			warnings: [],
		}
	}
	if (isDataFieldsOnlyColumnsPivot(pivot)) {
		const expected = aggregateDataFieldsOnlyColumnsPivotOutput(cache, pivot, cacheRows)
		if (!expected.ok) return unsupportedPivotAudit(base, expected.warning)
		const actual = readMultiRowAxisPivotOutput(
			workbook,
			sheet.id,
			pivot,
			expected.value.rowKeys,
			expected.value.columnKeys,
		)
		if (!actual.ok) return unsupportedPivotAudit(base, actual.warning)
		const mismatches = comparePivotOutput(expected.value.values, actual.value)
		return {
			...base,
			status: mismatches.length > 0 ? 'mismatch' : 'passed',
			checkedValueCount: expected.value.checkedValueCount,
			mismatches,
			warnings: [],
		}
	}
	if (isAxisItemMatrixPivot(pivot)) {
		const expected = aggregateMultiRowAxisPivotOutput(cache, pivot, cacheRows)
		if (!expected.ok) return unsupportedPivotAudit(base, expected.warning)
		const actual = readMultiRowAxisPivotOutput(
			workbook,
			sheet.id,
			pivot,
			expected.value.rowKeys,
			expected.value.columnKeys,
		)
		if (!actual.ok) return unsupportedPivotAudit(base, actual.warning)
		const mismatches = comparePivotOutput(expected.value.values, actual.value)
		return {
			...base,
			status: mismatches.length > 0 ? 'mismatch' : 'passed',
			checkedValueCount: expected.value.checkedValueCount,
			mismatches,
			warnings: [],
		}
	}
	if (isMultiRowAxisSingleDataPivot(pivot)) {
		const expected = aggregateMultiRowAxisSingleDataPivotOutput(cache, pivot, cacheRows)
		if (!expected.ok) return unsupportedPivotAudit(base, expected.warning)
		const actual = readMultiRowAxisPivotOutput(
			workbook,
			sheet.id,
			pivot,
			expected.value.rowKeys,
			expected.value.columnKeys,
		)
		if (!actual.ok) return unsupportedPivotAudit(base, actual.warning)
		const mismatches = comparePivotOutput(expected.value.values, actual.value)
		return {
			...base,
			status: mismatches.length > 0 ? 'mismatch' : 'passed',
			checkedValueCount: expected.value.checkedValueCount,
			mismatches,
			warnings: [],
		}
	}
	if (pivot.rowFields.length !== 1) {
		return unsupportedPivotAudit(base, 'Only one-row-field pivots are audited.')
	}
	if (pivot.columnFields.length > 0 && !pivot.columnFields.every((field) => field.index === -2)) {
		return unsupportedPivotAudit(
			base,
			'Column-field pivots beyond the data-field axis are not audited.',
		)
	}
	const rowFieldIndex = pivot.rowFields[0]?.index
	const rowField = rowFieldIndex === undefined ? undefined : cache.fields[rowFieldIndex]
	if (!rowField) return unsupportedPivotAudit(base, 'Pivot row field metadata was not found.')
	const expected = aggregateSimplePivotOutput(cache, pivot, rowField.index, cacheRows)
	if (!expected.ok) return unsupportedPivotAudit(base, expected.warning)
	const actual = readSimplePivotOutput(workbook, sheet.id, pivot, expected.value.dataFieldNames)
	if (!actual.ok) return unsupportedPivotAudit(base, actual.warning)
	const mismatches = comparePivotOutput(expected.value.values, actual.value)
	return {
		...base,
		status: mismatches.length > 0 ? 'mismatch' : 'passed',
		checkedValueCount: expected.value.checkedValueCount,
		mismatches,
		warnings: [],
	}
}

function pivotCacheWithMaterializedAuditRows(
	workbook: Workbook,
	cache: PivotCacheInfo,
): { ok: true; value: PivotCacheInfo } | { ok: false; warning: string } {
	if (cache.records?.materializedComplete && cache.records.materializedRecords) {
		return { ok: true, value: cache }
	}
	const sourceRecords = materializePivotCacheRecordsFromWorksheetSource(workbook, cache)
	if (!sourceRecords.ok) return sourceRecords
	return {
		ok: true,
		value: {
			...cache,
			records: {
				partPath:
					cache.records?.partPath ?? cache.recordsPartPath ?? `${cache.partPath}#worksheet-source`,
				...(cache.recordCount !== undefined ? { declaredCount: cache.recordCount } : {}),
				parsedCount: sourceRecords.value.length,
				preview: sourceRecords.value.slice(0, 5),
				materializedRecords: sourceRecords.value,
				materializedCount: sourceRecords.value.length,
				materializedComplete: true,
				valueKindCounts: pivotCacheRecordValueKindCounts(sourceRecords.value),
			},
		},
	}
}

function materializePivotCacheRecordsFromWorksheetSource(
	workbook: Workbook,
	cache: PivotCacheInfo,
): { ok: true; value: readonly PivotCacheRecordInfo[] } | { ok: false; warning: string } {
	if (cache.sourceType !== 'worksheet' || !cache.sourceSheet || !cache.sourceRef) {
		return { ok: false, warning: 'Pivot cache records are not fully materialized.' }
	}
	let sourceRange: RangeRef
	try {
		sourceRange = parseRange(cache.sourceRef)
	} catch {
		return { ok: false, warning: `Pivot cache source range is invalid: ${cache.sourceRef}` }
	}
	const sourceSheet = workbook.getSheet(cache.sourceSheet)
	if (!sourceSheet) return { ok: false, warning: 'Pivot cache source sheet was not loaded.' }
	const sourceFieldCount = sourceRange.end.col - sourceRange.start.col + 1
	if (sourceFieldCount !== cache.fields.length) {
		return { ok: false, warning: 'Pivot cache source range does not match cache fields.' }
	}
	for (let offset = 0; offset < sourceFieldCount; offset++) {
		const field = cache.fields[offset]
		const header = sourceSheet.cells.get(sourceRange.start.row, sourceRange.start.col + offset)
		if (field?.name !== pivotCacheSourceHeaderName(header?.value ?? EMPTY)) {
			return { ok: false, warning: 'Pivot cache source headers do not match cache fields.' }
		}
	}
	const records: PivotCacheRecordInfo[] = []
	for (let row = sourceRange.start.row + 1; row <= sourceRange.end.row; row++) {
		const values: PivotCacheRecordValueInfo[] = []
		for (let offset = 0; offset < sourceFieldCount; offset++) {
			const value = sourceSheet.cells.get(row, sourceRange.start.col + offset)?.value ?? EMPTY
			values.push(pivotCacheRecordValueFromCell(offset, value))
		}
		records.push({ index: records.length, values })
	}
	return { ok: true, value: records }
}

function pivotCacheSourceHeaderName(value: CellValue): string | undefined {
	switch (value.kind) {
		case 'string':
			return value.value
		case 'richText':
			return value.runs.map((run) => run.text).join('')
		case 'number':
			return String(value.value)
		case 'date':
			return String(value.serial)
		case 'boolean':
			return value.value ? 'TRUE' : 'FALSE'
		case 'error':
			return value.value
		default:
			return undefined
	}
}

function pivotCacheRecordValueFromCell(index: number, value: CellValue): PivotCacheRecordValueInfo {
	switch (value.kind) {
		case 'string':
			return { index, kind: 'string', value: value.value }
		case 'richText':
			return { index, kind: 'string', value: value.runs.map((run) => run.text).join('') }
		case 'number':
			return { index, kind: 'number', value: String(value.value) }
		case 'date':
			return { index, kind: 'date', value: String(value.serial) }
		case 'boolean':
			return { index, kind: 'boolean', value: value.value ? '1' : '0' }
		case 'error':
			return { index, kind: 'error', value: value.value }
		default:
			return { index, kind: 'missing' }
	}
}

function pivotCacheRecordValueKindCounts(
	records: readonly PivotCacheRecordInfo[],
): readonly PivotCacheRecordValueKindCount[] {
	const counts = new Map<PivotCacheRecordValueInfo['kind'], number>()
	for (const record of records) {
		for (const value of record.values) {
			counts.set(value.kind, (counts.get(value.kind) ?? 0) + 1)
		}
	}
	return [...counts.entries()].map(([kind, count]) => ({ kind, count }))
}

function isAxisItemMatrixPivot(pivot: PivotTableInfo): boolean {
	return (
		pivot.rowFields.length > 0 &&
		pivot.columnFields.length > 0 &&
		pivot.columnFields.every((field) => field.index >= 0)
	)
}

function isMultiRowAxisSingleDataPivot(pivot: PivotTableInfo): boolean {
	return (
		pivot.rowFields.length > 1 &&
		pivot.rowFields.every((field) => field.index >= 0) &&
		pivot.columnFields.length === 0 &&
		pivot.dataFields.length === 1
	)
}

interface PivotAxisOutputItem {
	readonly key: string
	readonly filters: PivotAuditFilters
}

const PIVOT_MISSING_ITEM_FILTER_VALUE = '__ASCEND_PIVOT_MISSING_ITEM__'

function isEmptyPivotOutput(pivot: PivotTableInfo): boolean {
	return (
		pivot.rowFields.length === 0 && pivot.columnFields.length === 0 && pivot.dataFields.length === 0
	)
}

function hasSavedSheetCells(sheet: Workbook['sheets'][number]): boolean {
	return sheet.cells.usedRange() !== null
}

function isDataFieldsOnRowsPivot(pivot: PivotTableInfo): boolean {
	return (
		pivot.rowFields.length === 1 &&
		pivot.rowFields[0]?.index === -2 &&
		pivot.options?.dataOnRows === true &&
		pivot.columnFields.length > 0
	)
}

function isDataFieldsNestedOnRowsPivot(pivot: PivotTableInfo): boolean {
	return (
		pivot.rowFields.length === 2 &&
		pivot.rowFields[0]?.index !== undefined &&
		pivot.rowFields[0].index >= 0 &&
		pivot.rowFields[1]?.index === -2 &&
		pivot.options?.dataOnRows === true &&
		pivot.dataFields.length > 1 &&
		pivot.pageFields.length === 0
	)
}

function isDataFieldsOnColumnsPivot(pivot: PivotTableInfo): boolean {
	return (
		pivot.rowFields.length > 0 &&
		pivot.columnFields.some((field) => field.index === -2) &&
		pivot.columnFields.some((field) => field.index >= 0) &&
		pivot.dataFields.length > 1
	)
}

function isDataFieldsOnlyColumnsPivot(pivot: PivotTableInfo): boolean {
	return (
		pivot.rowFields.length > 1 &&
		pivot.rowFields.every((field) => field.index >= 0) &&
		pivot.columnFields.length > 0 &&
		pivot.columnFields.every((field) => field.index === -2) &&
		pivot.dataFields.length > 1
	)
}

function unsupportedPivotAudit(
	base: Pick<PivotOutputAuditInfo, 'partPath' | 'sheetName' | 'pivotTable' | 'cacheId'>,
	warning: string,
): PivotOutputAuditInfo {
	return {
		...base,
		status: 'unsupported',
		checkedValueCount: 0,
		mismatches: [],
		warnings: [warning],
	}
}

function auditEmptyPivotOutput(
	workbook: Workbook,
	sheetId: string,
	pivot: PivotTableInfo,
): { ok: true; checkedCellCount: number } | { ok: false; warning: string } {
	if (!pivot.locationRef) return { ok: false, warning: 'Pivot table has no output location.' }
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return { ok: false, warning: `Pivot output range is invalid: ${pivot.locationRef}` }
	}
	const sheet = workbook.sheets.find((entry) => entry.id === sheetId)
	if (!sheet) return { ok: false, warning: 'Pivot output sheet was not loaded.' }
	let checkedCellCount = 0
	for (let row = bounds.start.row; row <= bounds.end.row; row++) {
		for (let col = bounds.start.col; col <= bounds.end.col; col++) {
			checkedCellCount++
			if ((sheet.cells.get(row, col)?.value ?? EMPTY).kind !== 'empty') {
				return { ok: false, warning: 'Empty pivot output range contains saved cell values.' }
			}
		}
	}
	return { ok: true, checkedCellCount }
}

function auditRowLabelOnlyPivotOutput(
	workbook: Workbook,
	sheetId: string,
	pivot: PivotTableInfo,
	cache: PivotCacheInfo,
): { ok: true; checkedCellCount: number } | { ok: false; warning: string } {
	if (pivot.columnFields.length > 0) {
		return { ok: false, warning: 'No-data pivot column labels are not audited.' }
	}
	if (pivot.rowFields.length !== 1) {
		return { ok: false, warning: 'No-data pivot audits require exactly one row field.' }
	}
	if (!pivot.locationRef) return { ok: false, warning: 'Pivot table has no output location.' }
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return { ok: false, warning: `Pivot output range is invalid: ${pivot.locationRef}` }
	}
	const sheet = workbook.sheets.find((entry) => entry.id === sheetId)
	if (!sheet) return { ok: false, warning: 'Pivot output sheet was not loaded.' }
	const rowFieldIndex = pivot.rowFields[0]?.index
	if (rowFieldIndex === undefined || rowFieldIndex < 0) {
		return { ok: false, warning: 'No-data pivot row field metadata was not found.' }
	}
	const rows = buildPivotSimpleRowOutputItems(cache, pivot, rowFieldIndex)
	if (!rows.ok) return rows
	const dataStartOffset = pivot.location?.firstDataRow ?? 1
	const dataStartRow = bounds.start.row + dataStartOffset
	const labelCol = bounds.start.col
	if (dataStartRow + rows.value.length - 1 > bounds.end.row) {
		return { ok: false, warning: 'No-data pivot output range is missing saved row labels.' }
	}
	for (let index = 0; index < rows.value.length; index++) {
		const expected = normalizeSavedPivotLabel(rows.value[index]?.key ?? '')
		const actual = normalizeSavedPivotLabel(
			cellText(sheet.cells.get(dataStartRow + index, labelCol)?.value ?? EMPTY),
		)
		if (actual !== expected) {
			return { ok: false, warning: 'No-data pivot output labels did not match saved cells.' }
		}
	}
	return { ok: true, checkedCellCount: rows.value.length }
}

function aggregateDataFieldsOnRowsPivotOutput(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	cacheRows: readonly PivotCacheMaterializedRowInfo[],
):
	| {
			ok: true
			value: {
				dataFieldNames: readonly string[]
				columnKeys: readonly string[]
				checkedValueCount: number
				values: ReadonlyMap<string, ReadonlyMap<string, number>>
			}
	  }
	| { ok: false; warning: string } {
	if (!pivot.columnFields.every((field) => field.index >= 0)) {
		return { ok: false, warning: 'Data-field axis columns are not audited in row-data pivots.' }
	}
	const pageFilters = buildSimplePivotPageFilters(cache, pivot)
	if (!pageFilters.ok) return pageFilters
	const columns = buildPivotColumnOutputColumns(cache, pivot)
	if (!columns.ok) return columns
	const dataFieldNames = pivot.dataFields.map((field, index) =>
		pivotDataFieldAuditName(cache, field, index),
	)
	const output = new Map<string, Map<string, number>>()
	for (const dataFieldName of dataFieldNames) {
		const byColumn = new Map<string, number>()
		for (const column of columns.value) byColumn.set(column.key, 0)
		output.set(dataFieldName, byColumn)
	}
	for (const row of cacheRows) {
		if (!pivotCacheRowMatchesFilters(row, pageFilters.value)) continue
		for (const column of columns.value) {
			if (!column.matches(row)) continue
			for (let i = 0; i < pivot.dataFields.length; i++) {
				const dataField = pivot.dataFields[i]
				const dataFieldName = dataFieldNames[i] ?? `DataField${i + 1}`
				if (!dataField) continue
				const measured = measurePivotDataField(cache, row, dataField)
				if (!measured.ok) return measured
				addPivotOutput(output, dataFieldName, column.key, measured.value)
			}
		}
	}
	return {
		ok: true,
		value: {
			dataFieldNames,
			columnKeys: columns.value.map((column) => column.key),
			checkedValueCount: dataFieldNames.length * columns.value.length,
			values: output,
		},
	}
}

function buildPivotColumnOutputColumns(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
):
	| {
			ok: true
			value: readonly {
				key: string
				matches(row: PivotCacheMaterializedRowInfo): boolean
			}[]
	  }
	| { ok: false; warning: string } {
	if (!pivot.columnItems || pivot.columnItems.length === 0) {
		return { ok: false, warning: 'Pivot column output items were not found.' }
	}
	const columns: {
		key: string
		matches(row: PivotCacheMaterializedRowInfo): boolean
	}[] = []
	for (const item of pivot.columnItems) {
		const key = item.itemType === 'grand' ? 'Grand Total' : `Column ${item.index + 1}`
		if (item.itemType === 'grand') {
			columns.push({ key, matches: () => true })
			continue
		}
		const filters = new Map<number, Set<string>>()
		for (const fieldItem of item.fieldItems) {
			const axisField = pivot.columnFields[fieldItem.index]
			if (!axisField || axisField.index < 0) {
				return { ok: false, warning: 'Pivot column field item axis was not resolved.' }
			}
			if (fieldItem.item === undefined) {
				return { ok: false, warning: 'Repeated pivot column field items are not audited.' }
			}
			const pivotField = pivot.fields[axisField.index]
			const pivotItem = pivotField?.items?.[fieldItem.item]
			if (
				!pivotItem ||
				pivotItem.hidden ||
				pivotItem.missing ||
				pivotItem.cacheIndex === undefined
			) {
				return { ok: false, warning: 'Pivot column field item cache value was not resolved.' }
			}
			const sharedItem = cache.fields[axisField.index]?.sharedItems?.find(
				(entry) => entry.index === pivotItem.cacheIndex,
			)
			if (sharedItem?.value === undefined) {
				return { ok: false, warning: 'Pivot column shared item value was not resolved.' }
			}
			filters.set(axisField.index, new Set([sharedItem.value]))
		}
		if (filters.size !== pivot.columnFields.length) {
			return { ok: false, warning: 'Pivot column item does not cover every column field.' }
		}
		columns.push({
			key,
			matches: (row) => pivotCacheRowMatchesFilters(row, filters),
		})
	}
	return { ok: true, value: columns }
}

interface PivotDataFieldColumnOutputItem extends PivotAxisOutputItem {
	readonly dataFieldIndex: number
	readonly dataField: PivotTableInfo['dataFields'][number]
	readonly isGrand: boolean
}

interface PivotDataFieldRowOutputItem extends PivotAxisOutputItem {
	readonly dataFieldIndex: number
	readonly dataField: PivotTableInfo['dataFields'][number]
}

function aggregateDataFieldsNestedOnRowsPivotOutput(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	cacheRows: readonly PivotCacheMaterializedRowInfo[],
):
	| {
			ok: true
			value: {
				rowKeys: readonly string[]
				columnKeys: readonly string[]
				checkedValueCount: number
				values: ReadonlyMap<string, ReadonlyMap<string, number>>
			}
	  }
	| { ok: false; warning: string } {
	if (!pivot.columnFields.every((field) => field.index >= 0)) {
		return { ok: false, warning: 'Nested row data-field pivots require real column fields.' }
	}
	const pageFilters = buildSimplePivotPageFilters(cache, pivot)
	if (!pageFilters.ok) return pageFilters
	const rows = buildPivotDataFieldRowOutputItems(cache, pivot)
	if (!rows.ok) return rows
	const columns =
		pivot.columnFields.length > 0 ? buildPivotAxisOutputItems(cache, pivot, 'column') : undefined
	if (columns && !columns.ok) return columns
	const columnItems = columns?.value ?? [{ key: 'Total', filters: new Map<number, Set<string>>() }]
	const output = new Map<string, Map<string, number>>()
	for (const row of rows.value) {
		const byColumn = new Map<string, number>()
		for (const column of columnItems) byColumn.set(column.key, 0)
		output.set(row.key, byColumn)
	}
	const rowMatcher = buildPivotAxisOutputMatcher(rows.value)
	const columnMatcher = buildPivotAxisOutputMatcher(columnItems)
	for (const cacheRow of cacheRows) {
		if (!pivotCacheRowMatchesFilters(cacheRow, pageFilters.value)) continue
		for (const rowItem of matchingPivotAxisOutputItems(cacheRow, rowMatcher)) {
			const measured = measurePivotDataField(cache, cacheRow, rowItem.dataField)
			if (!measured.ok) return measured
			for (const columnItem of matchingPivotAxisOutputItems(cacheRow, columnMatcher)) {
				addPivotOutput(output, rowItem.key, columnItem.key, measured.value)
			}
		}
	}
	return {
		ok: true,
		value: {
			rowKeys: rows.value.map((row) => row.key),
			columnKeys: columnItems.map((column) => column.key),
			checkedValueCount: rows.value.length * columnItems.length,
			values: output,
		},
	}
}

function buildPivotDataFieldRowOutputItems(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
): { ok: true; value: readonly PivotDataFieldRowOutputItem[] } | { ok: false; warning: string } {
	if (!pivot.rowItems || pivot.rowItems.length === 0) {
		return { ok: false, warning: 'Pivot row output items were not found.' }
	}
	const rows: PivotDataFieldRowOutputItem[] = []
	const previousItems: Array<number | undefined> = []
	const seenLabels = new Map<string, number>()
	for (const item of pivot.rowItems) {
		const filters = new Map<number, Set<string>>()
		let dataFieldIndex = item.dataFieldIndex
		let label: string | undefined
		if (item.itemType === 'grand') {
			const grandFilters = buildPivotDataFieldAxisGrandFilters(cache, pivot, pivot.rowFields)
			if (!grandFilters.ok) return grandFilters
			addPivotAxisFilters(filters, grandFilters.value)
			dataFieldIndex ??= resolvePivotRowDataFieldAxisItem(pivot, item, previousItems)
			const dataField = pivot.dataFields[dataFieldIndex]
			if (!dataField) return { ok: false, warning: 'Pivot row data field was not resolved.' }
			rows.push({
				key: uniquePivotAxisKey(
					pivotDataFieldRowKey(cache, dataField, dataFieldIndex, 'Grand Total'),
					seenLabels,
				),
				filters,
				dataFieldIndex,
				dataField,
			})
			continue
		}
		const repeated = item.repeatedItemCount ?? 0
		for (let position = 0; position < repeated && position < pivot.rowFields.length; position++) {
			const axisField = pivot.rowFields[position]
			const previousItem = previousItems[position]
			if (!axisField || previousItem === undefined) continue
			if (axisField.index === -2) {
				dataFieldIndex ??= previousItem
				continue
			}
			const resolved = resolvePivotAxisFieldItemFilter(cache, pivot, axisField.index, previousItem)
			if (!resolved.ok) return resolved
			addPivotAxisFilters(filters, resolved.value.filters)
			label = resolved.value.label
		}
		for (const fieldItem of item.fieldItems) {
			const position = repeated + fieldItem.index
			const axisField = pivot.rowFields[position]
			if (!axisField) return { ok: false, warning: 'Pivot row field item axis was not resolved.' }
			const itemIndex = fieldItem.item ?? 0
			previousItems[position] = itemIndex
			if (axisField.index === -2) {
				dataFieldIndex = itemIndex
				continue
			}
			const resolved = resolvePivotAxisFieldItemFilter(cache, pivot, axisField.index, itemIndex)
			if (!resolved.ok) return resolved
			addPivotAxisFilters(filters, resolved.value.filters)
			label = resolved.value.label
		}
		dataFieldIndex ??= 0
		const dataField = pivot.dataFields[dataFieldIndex]
		if (!dataField) return { ok: false, warning: 'Pivot row data field was not resolved.' }
		if (!label) return { ok: false, warning: 'Pivot row output item label was not resolved.' }
		rows.push({
			key: uniquePivotAxisKey(
				pivotDataFieldRowKey(cache, dataField, dataFieldIndex, label),
				seenLabels,
			),
			filters,
			dataFieldIndex,
			dataField,
		})
	}
	return { ok: true, value: rows }
}

function resolvePivotRowDataFieldAxisItem(
	pivot: PivotTableInfo,
	item: NonNullable<PivotTableInfo['rowItems']>[number],
	previousItems: readonly (number | undefined)[],
): number {
	for (const fieldItem of item.fieldItems) {
		const axisField = pivot.rowFields[fieldItem.index]
		if (axisField?.index === -2) return fieldItem.item ?? 0
	}
	if (item.itemType === 'grand') return 0
	const dataAxisPosition = pivot.rowFields.findIndex((field) => field.index === -2)
	return dataAxisPosition >= 0 ? (previousItems[dataAxisPosition] ?? 0) : 0
}

function pivotDataFieldRowKey(
	cache: PivotCacheInfo,
	dataField: PivotTableInfo['dataFields'][number],
	dataFieldIndex: number,
	label: string,
): string {
	return `${label} / ${pivotDataFieldAuditName(cache, dataField, dataFieldIndex)}`
}

function aggregateDataFieldsOnColumnsPivotOutput(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	cacheRows: readonly PivotCacheMaterializedRowInfo[],
):
	| {
			ok: true
			value: {
				rowKeys: readonly string[]
				columnKeys: readonly string[]
				checkedValueCount: number
				values: ReadonlyMap<string, ReadonlyMap<string, number>>
			}
	  }
	| { ok: false; warning: string } {
	if (!pivot.columnFields.every((field) => field.index === -2 || field.index >= 0)) {
		return { ok: false, warning: 'Unsupported negative pivot column field axis.' }
	}
	const pageFilters = buildSimplePivotPageFilters(cache, pivot)
	if (!pageFilters.ok) return pageFilters
	const rows = buildPivotAxisOutputItems(cache, pivot, 'row')
	if (!rows.ok) return rows
	const columns = buildPivotDataFieldColumnOutputItems(cache, pivot)
	if (!columns.ok) return columns
	const output = new Map<string, Map<string, number>>()
	for (const row of rows.value) {
		const byColumn = new Map<string, number>()
		for (const column of columns.value) byColumn.set(column.key, 0)
		output.set(row.key, byColumn)
	}
	const rowMatcher = buildPivotAxisOutputMatcher(rows.value)
	const columnMatcher = buildPivotAxisOutputMatcher(columns.value)
	for (const row of cacheRows) {
		if (!pivotCacheRowMatchesFilters(row, pageFilters.value)) continue
		for (const rowItem of matchingPivotAxisOutputItems(row, rowMatcher)) {
			for (const columnItem of matchingPivotAxisOutputItems(row, columnMatcher)) {
				const measured = measurePivotDataField(cache, row, columnItem.dataField)
				if (!measured.ok) return measured
				addPivotOutput(output, rowItem.key, columnItem.key, measured.value)
			}
		}
	}
	const showDataAs = applyPivotDataFieldColumnShowDataAs(output, columns.value)
	if (!showDataAs.ok) return showDataAs
	return {
		ok: true,
		value: {
			rowKeys: rows.value.map((row) => row.key),
			columnKeys: columns.value.map((column) => column.key),
			checkedValueCount: rows.value.length * columns.value.length,
			values: showDataAs.value,
		},
	}
}

function buildPivotDataFieldColumnOutputItems(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
): { ok: true; value: readonly PivotDataFieldColumnOutputItem[] } | { ok: false; warning: string } {
	if (!pivot.columnItems || pivot.columnItems.length === 0) {
		return { ok: false, warning: 'Pivot column output items were not found.' }
	}
	const columns: PivotDataFieldColumnOutputItem[] = []
	const previousItems: Array<number | undefined> = []
	const seenLabels = new Map<string, number>()
	for (const item of pivot.columnItems) {
		const filters = new Map<number, Set<string>>()
		let dataFieldIndex = item.dataFieldIndex
		let label: string | undefined
		if (item.itemType === 'grand') {
			const grandFilters = buildPivotColumnGrandFilters(cache, pivot)
			if (!grandFilters.ok) return grandFilters
			addPivotAxisFilters(filters, grandFilters.value)
			dataFieldIndex ??= resolvePivotDataFieldAxisItem(pivot, item, previousItems)
			const dataField = pivot.dataFields[dataFieldIndex]
			if (!dataField) return { ok: false, warning: 'Pivot column data field was not resolved.' }
			columns.push({
				key: uniquePivotAxisKey(
					pivotDataFieldColumnKey(cache, dataField, dataFieldIndex, 'Grand Total'),
					seenLabels,
				),
				filters,
				dataFieldIndex,
				dataField,
				isGrand: true,
			})
			continue
		}
		const repeated = item.repeatedItemCount ?? 0
		for (
			let position = 0;
			position < repeated && position < pivot.columnFields.length;
			position++
		) {
			const axisField = pivot.columnFields[position]
			const previousItem = previousItems[position]
			if (!axisField || previousItem === undefined) continue
			if (axisField.index === -2) {
				dataFieldIndex ??= previousItem
				continue
			}
			const resolved = resolvePivotAxisFieldItemFilter(cache, pivot, axisField.index, previousItem)
			if (!resolved.ok) return resolved
			addPivotAxisFilters(filters, resolved.value.filters)
			label = resolved.value.label
		}
		for (const fieldItem of item.fieldItems) {
			const position = repeated + fieldItem.index
			const axisField = pivot.columnFields[position]
			if (!axisField)
				return { ok: false, warning: 'Pivot column field item axis was not resolved.' }
			const itemIndex = fieldItem.item ?? 0
			previousItems[position] = itemIndex
			if (axisField.index === -2) {
				dataFieldIndex = itemIndex
				continue
			}
			const resolved = resolvePivotAxisFieldItemFilter(cache, pivot, axisField.index, itemIndex)
			if (!resolved.ok) return resolved
			addPivotAxisFilters(filters, resolved.value.filters)
			label = resolved.value.label
		}
		dataFieldIndex ??= 0
		const dataField = pivot.dataFields[dataFieldIndex]
		if (!dataField) return { ok: false, warning: 'Pivot column data field was not resolved.' }
		if (!label) return { ok: false, warning: 'Pivot column output item label was not resolved.' }
		columns.push({
			key: uniquePivotAxisKey(
				pivotDataFieldColumnKey(cache, dataField, dataFieldIndex, label),
				seenLabels,
			),
			filters,
			dataFieldIndex,
			dataField,
			isGrand: false,
		})
	}
	return { ok: true, value: columns }
}

function buildPivotDataFieldOnlyColumnOutputItems(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
): { ok: true; value: readonly PivotDataFieldColumnOutputItem[] } | { ok: false; warning: string } {
	if (!pivot.columnFields.every((field) => field.index === -2)) {
		return { ok: false, warning: 'Pivot data-field-only columns had non-data axis fields.' }
	}
	const seenLabels = new Map<string, number>()
	const columns: PivotDataFieldColumnOutputItem[] = []
	const items =
		pivot.columnItems && pivot.columnItems.length > 0
			? pivot.columnItems
			: pivot.dataFields.map((_, index) => ({
					index,
					fieldItems: [{ index: 0, item: index }],
					dataFieldIndex: index,
				}))
	for (const item of items) {
		let dataFieldIndex = item.dataFieldIndex
		for (const fieldItem of item.fieldItems) {
			const axisField = pivot.columnFields[fieldItem.index]
			if (axisField?.index !== -2) {
				return { ok: false, warning: 'Pivot column data field axis was not resolved.' }
			}
			dataFieldIndex ??= fieldItem.item ?? 0
		}
		dataFieldIndex ??= 0
		const dataField = pivot.dataFields[dataFieldIndex]
		if (!dataField) return { ok: false, warning: 'Pivot column data field was not resolved.' }
		columns.push({
			key: uniquePivotAxisKey(
				pivotDataFieldAuditName(cache, dataField, dataFieldIndex),
				seenLabels,
			),
			filters: new Map(),
			dataFieldIndex,
			dataField,
			isGrand: false,
		})
	}
	return columns.length > 0
		? { ok: true, value: columns }
		: { ok: false, warning: 'Pivot data-field-only column items were not found.' }
}

function resolvePivotDataFieldAxisItem(
	pivot: PivotTableInfo,
	item: NonNullable<PivotTableInfo['columnItems']>[number],
	previousItems: readonly (number | undefined)[],
): number {
	for (const fieldItem of item.fieldItems) {
		const axisField = pivot.columnFields[fieldItem.index]
		if (axisField?.index === -2) return fieldItem.item ?? 0
	}
	const dataAxisPosition = pivot.columnFields.findIndex((field) => field.index === -2)
	return dataAxisPosition >= 0 ? (previousItems[dataAxisPosition] ?? 0) : 0
}

function buildPivotDataFieldAxisGrandFilters(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	fields: readonly { readonly index: number }[],
): { ok: true; value: PivotAuditFilters } | { ok: false; warning: string } {
	const filters = new Map<number, Set<string>>()
	for (const axisField of fields) {
		if (axisField.index === -2) continue
		if (axisField.index < 0) return { ok: false, warning: 'Unsupported negative pivot axis field.' }
		const values = visiblePivotAxisFieldValues(cache, pivot, axisField.index)
		if (!values.ok) return values
		if (values.value.fieldIndex !== undefined && values.value.values.size > 0) {
			filters.set(values.value.fieldIndex, values.value.values)
		}
	}
	return { ok: true, value: filters }
}

function buildPivotColumnGrandFilters(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
): { ok: true; value: PivotAuditFilters } | { ok: false; warning: string } {
	return buildPivotDataFieldAxisGrandFilters(cache, pivot, pivot.columnFields)
}

function pivotDataFieldColumnKey(
	cache: PivotCacheInfo,
	dataField: PivotTableInfo['dataFields'][number],
	dataFieldIndex: number,
	label: string,
): string {
	return `${pivotDataFieldAuditName(cache, dataField, dataFieldIndex)} / ${label}`
}

function pivotDataFieldAuditName(
	cache: PivotCacheInfo,
	dataField: PivotTableInfo['dataFields'][number],
	dataFieldIndex: number,
): string {
	if (dataField.name) return dataField.name
	const fieldName = cache.fields[dataField.fieldIndex]?.name
	const subtotal = pivotDataFieldSubtotalLabel(dataField.subtotal)
	return fieldName ? `${subtotal} - ${fieldName}` : `DataField${dataFieldIndex + 1}`
}

function pivotDataFieldSubtotalLabel(
	subtotal: PivotTableInfo['dataFields'][number]['subtotal'],
): string {
	switch (subtotal) {
		case 'count':
			return 'Count'
		case 'average':
			return 'Average'
		default:
			return 'Sum'
	}
}

function applyPivotDataFieldColumnShowDataAs(
	output: ReadonlyMap<string, ReadonlyMap<string, number>>,
	columns: readonly PivotDataFieldColumnOutputItem[],
):
	| { ok: true; value: ReadonlyMap<string, ReadonlyMap<string, number>> }
	| { ok: false; warning: string } {
	const transformed = new Map<string, Map<string, number>>()
	for (const [rowKey, valuesByColumn] of output) {
		transformed.set(rowKey, new Map(valuesByColumn))
	}
	const dataFieldIndexes = new Set(columns.map((column) => column.dataFieldIndex))
	for (const dataFieldIndex of dataFieldIndexes) {
		const dataField = columns.find((column) => column.dataFieldIndex === dataFieldIndex)?.dataField
		if (!dataField?.showDataAs) continue
		if (dataField.showDataAs !== 'percentOfRow') {
			return { ok: false, warning: 'Pivot show-data-as calculations are not audited.' }
		}
		const fieldColumns = columns.filter((column) => column.dataFieldIndex === dataFieldIndex)
		const grandColumn = fieldColumns.find((column) => column.isGrand)
		for (const [rowKey, valuesByColumn] of output) {
			const rowTotal =
				(grandColumn ? valuesByColumn.get(grandColumn.key) : undefined) ??
				fieldColumns.reduce(
					(total, column) =>
						column.isGrand ? total : total + (valuesByColumn.get(column.key) ?? 0),
					0,
				)
			const transformedRow = transformed.get(rowKey)
			if (!transformedRow) continue
			for (const column of fieldColumns) {
				const raw = valuesByColumn.get(column.key) ?? 0
				transformedRow.set(column.key, rowTotal === 0 ? 0 : raw / rowTotal)
			}
		}
	}
	return { ok: true, value: transformed }
}

function aggregateMultiRowAxisPivotOutput(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	cacheRows: readonly PivotCacheMaterializedRowInfo[],
):
	| {
			ok: true
			value: {
				rowKeys: readonly string[]
				columnKeys: readonly string[]
				checkedValueCount: number
				values: ReadonlyMap<string, ReadonlyMap<string, number>>
			}
	  }
	| { ok: false; warning: string } {
	if (pivot.dataFields.length !== 1) {
		return { ok: false, warning: 'Multi-row pivot audits require exactly one data field.' }
	}
	if (pivot.columnFields.length === 0) {
		return { ok: false, warning: 'Multi-row pivot audits require column output items.' }
	}
	if (!pivot.columnFields.every((field) => field.index >= 0)) {
		return { ok: false, warning: 'Multi-row pivot data-field axis columns are not audited.' }
	}
	const pageFilters = buildSimplePivotPageFilters(cache, pivot)
	if (!pageFilters.ok) return pageFilters
	const rows = buildPivotAxisOutputItemsWithSingleFieldFallback(cache, pivot, 'row')
	if (!rows.ok) return rows
	const columns = buildPivotAxisOutputItemsWithTwoFieldSubtotalFallback(
		cache,
		pivot,
		'column',
		rows.value,
	)
	if (!columns.ok) return columns
	const dataField = pivot.dataFields[0]
	if (!dataField) return { ok: false, warning: 'Pivot data field metadata was not found.' }
	const output = new Map<string, Map<string, number>>()
	for (const row of rows.value) {
		const byColumn = new Map<string, number>()
		for (const column of columns.value) byColumn.set(column.key, 0)
		output.set(row.key, byColumn)
	}
	const rowMatcher = buildPivotAxisOutputMatcher(rows.value)
	const columnMatcher = buildPivotAxisOutputMatcher(columns.value)
	for (const row of cacheRows) {
		if (!pivotCacheRowMatchesFilters(row, pageFilters.value)) continue
		const measured = measurePivotDataField(cache, row, dataField)
		if (!measured.ok) return measured
		for (const rowItem of matchingPivotAxisOutputItems(row, rowMatcher)) {
			for (const columnItem of matchingPivotAxisOutputItems(row, columnMatcher)) {
				addPivotOutput(output, rowItem.key, columnItem.key, measured.value)
			}
		}
	}
	const showDataAs = applyPivotMatrixShowDataAs(
		output,
		dataField,
		columns.value.map((column) => column.key),
	)
	if (!showDataAs.ok) return showDataAs
	return {
		ok: true,
		value: {
			rowKeys: rows.value.map((row) => row.key),
			columnKeys: columns.value.map((column) => column.key),
			checkedValueCount: rows.value.length * columns.value.length,
			values: showDataAs.value,
		},
	}
}

function aggregateDataFieldsOnlyColumnsPivotOutput(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	cacheRows: readonly PivotCacheMaterializedRowInfo[],
):
	| {
			ok: true
			value: {
				rowKeys: readonly string[]
				columnKeys: readonly string[]
				checkedValueCount: number
				values: ReadonlyMap<string, ReadonlyMap<string, number>>
			}
	  }
	| { ok: false; warning: string } {
	const pageFilters = buildSimplePivotPageFilters(cache, pivot)
	if (!pageFilters.ok) return pageFilters
	const rows = buildPivotAxisOutputItems(cache, pivot, 'row')
	if (!rows.ok) return rows
	const columns = buildPivotDataFieldOnlyColumnOutputItems(cache, pivot)
	if (!columns.ok) return columns
	const output = new Map<string, Map<string, number>>()
	for (const row of rows.value) {
		const byColumn = new Map<string, number>()
		for (const column of columns.value) byColumn.set(column.key, 0)
		output.set(row.key, byColumn)
	}
	const rowMatcher = buildPivotAxisOutputMatcher(rows.value)
	for (const row of cacheRows) {
		if (!pivotCacheRowMatchesFilters(row, pageFilters.value)) continue
		for (const rowItem of matchingPivotAxisOutputItems(row, rowMatcher)) {
			for (const columnItem of columns.value) {
				const measured = measurePivotDataField(cache, row, columnItem.dataField)
				if (!measured.ok) return measured
				addPivotOutput(output, rowItem.key, columnItem.key, measured.value)
			}
		}
	}
	const showDataAs = applyPivotDataFieldColumnShowDataAs(output, columns.value)
	if (!showDataAs.ok) return showDataAs
	return {
		ok: true,
		value: {
			rowKeys: rows.value.map((row) => row.key),
			columnKeys: columns.value.map((column) => column.key),
			checkedValueCount: rows.value.length * columns.value.length,
			values: showDataAs.value,
		},
	}
}

function aggregateMultiRowAxisSingleDataPivotOutput(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	cacheRows: readonly PivotCacheMaterializedRowInfo[],
):
	| {
			ok: true
			value: {
				rowKeys: readonly string[]
				columnKeys: readonly string[]
				checkedValueCount: number
				values: ReadonlyMap<string, ReadonlyMap<string, number>>
			}
	  }
	| { ok: false; warning: string } {
	const pageFilters = buildSimplePivotPageFilters(cache, pivot)
	if (!pageFilters.ok) return pageFilters
	const rows = buildPivotAxisOutputItemsWithCompactTwoFieldFallback(cache, pivot, 'row')
	if (!rows.ok) return rows
	const dataField = pivot.dataFields[0]
	if (!dataField) return { ok: false, warning: 'Pivot data field metadata was not found.' }
	const columnKey = pivotDataFieldAuditName(cache, dataField, 0)
	const columnKeys = [columnKey]
	const output = new Map<string, Map<string, number>>()
	for (const row of rows.value) {
		output.set(row.key, new Map([[columnKey, 0]]))
	}
	const rowMatcher = buildPivotAxisOutputMatcher(rows.value)
	for (const row of cacheRows) {
		if (!pivotCacheRowMatchesFilters(row, pageFilters.value)) continue
		const measured = measurePivotDataField(cache, row, dataField)
		if (!measured.ok) return measured
		for (const rowItem of matchingPivotAxisOutputItems(row, rowMatcher)) {
			addPivotOutput(output, rowItem.key, columnKey, measured.value)
		}
	}
	const showDataAs = applyPivotMatrixShowDataAs(output, dataField, columnKeys)
	if (!showDataAs.ok) return showDataAs
	return {
		ok: true,
		value: {
			rowKeys: rows.value.map((row) => row.key),
			columnKeys,
			checkedValueCount: rows.value.length,
			values: showDataAs.value,
		},
	}
}

function applyPivotMatrixShowDataAs(
	output: ReadonlyMap<string, ReadonlyMap<string, number>>,
	dataField: PivotTableInfo['dataFields'][number],
	columnKeys: readonly string[],
):
	| { ok: true; value: ReadonlyMap<string, ReadonlyMap<string, number>> }
	| { ok: false; warning: string } {
	if (dataField.showDataAs === undefined) return { ok: true, value: output }
	if (dataField.showDataAs !== 'percentOfRow') {
		return { ok: false, warning: 'Pivot show-data-as calculations are not audited.' }
	}
	const transformed = new Map<string, Map<string, number>>()
	for (const [rowKey, valuesByColumn] of output) {
		const rowTotal =
			valuesByColumn.get('Grand Total') ??
			columnKeys.reduce(
				(total, columnKey) =>
					columnKey === 'Grand Total' ? total : total + (valuesByColumn.get(columnKey) ?? 0),
				0,
			)
		const transformedRow = new Map<string, number>()
		for (const columnKey of columnKeys) {
			const raw = valuesByColumn.get(columnKey) ?? 0
			transformedRow.set(columnKey, rowTotal === 0 ? 0 : raw / rowTotal)
		}
		transformed.set(rowKey, transformedRow)
	}
	return { ok: true, value: transformed }
}

function buildPivotAxisOutputItems(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	axis: 'row' | 'column',
): { ok: true; value: readonly PivotAxisOutputItem[] } | { ok: false; warning: string } {
	const fields = axis === 'row' ? pivot.rowFields : pivot.columnFields
	const items = axis === 'row' ? pivot.rowItems : pivot.columnItems
	if (!items || items.length === 0) {
		return { ok: false, warning: `Pivot ${axis} output items were not found.` }
	}
	const output: PivotAxisOutputItem[] = []
	const previousItems: Array<number | undefined> = []
	const seenLabels = new Map<string, number>()
	for (const item of items) {
		if (item.itemType === 'grand') {
			const filters = buildPivotAxisGrandFilters(cache, pivot, fields)
			if (!filters.ok) return filters
			output.push({ key: uniquePivotAxisKey('Grand Total', seenLabels), filters: filters.value })
			continue
		}
		const filters = new Map<number, Set<string>>()
		const repeated = item.repeatedItemCount ?? 0
		for (let position = 0; position < repeated && position < fields.length; position++) {
			const previousItem = previousItems[position]
			if (previousItem === undefined) continue
			const axisField = fields[position]
			if (!axisField) continue
			const resolved = resolvePivotAxisFieldItemFilter(cache, pivot, axisField.index, previousItem)
			if (!resolved.ok) return resolved
			addPivotAxisFilters(filters, resolved.value.filters)
		}
		let label: string | undefined
		for (const fieldItem of item.fieldItems) {
			const position = repeated + fieldItem.index
			const axisField = fields[position]
			if (!axisField)
				return { ok: false, warning: `Pivot ${axis} field item axis was not resolved.` }
			const itemIndex = fieldItem.item ?? 0
			previousItems[position] = itemIndex
			const resolved = resolvePivotAxisFieldItemFilter(cache, pivot, axisField.index, itemIndex)
			if (!resolved.ok) return resolved
			addPivotAxisFilters(filters, resolved.value.filters)
			label = resolved.value.label
		}
		if (!label) return { ok: false, warning: `Pivot ${axis} output item label was not resolved.` }
		output.push({ key: uniquePivotAxisKey(label, seenLabels), filters })
	}
	return { ok: true, value: output }
}

function buildPivotAxisOutputItemsWithSingleFieldFallback(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	axis: 'row' | 'column',
): { ok: true; value: readonly PivotAxisOutputItem[] } | { ok: false; warning: string } {
	const items = buildPivotAxisOutputItems(cache, pivot, axis)
	const missingWarning = `Pivot ${axis} output items were not found.`
	if (items.ok || items.warning !== missingWarning) return items
	return buildSingleFieldPivotAxisOutputItems(cache, pivot, axis, missingWarning)
}

function buildPivotAxisOutputItemsWithCompactTwoFieldFallback(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	axis: 'row' | 'column',
): { ok: true; value: readonly PivotAxisOutputItem[] } | { ok: false; warning: string } {
	const items = buildPivotAxisOutputItems(cache, pivot, axis)
	const missingWarning = `Pivot ${axis} output items were not found.`
	if (items.ok || items.warning !== missingWarning || axis !== 'row') return items
	return buildCompactTwoFieldRowAxisOutputItems(cache, pivot, missingWarning)
}

function buildPivotAxisOutputItemsWithTwoFieldSubtotalFallback(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	axis: 'row' | 'column',
	crossAxisItems?: readonly PivotAxisOutputItem[],
): { ok: true; value: readonly PivotAxisOutputItem[] } | { ok: false; warning: string } {
	const items = buildPivotAxisOutputItemsWithSingleFieldFallback(cache, pivot, axis)
	const missingWarning = `Pivot ${axis} output items were not found.`
	if (items.ok || items.warning !== missingWarning || axis !== 'column') return items
	return buildTwoFieldSubtotalAxisOutputItems(cache, pivot, axis, missingWarning, crossAxisItems)
}

function buildSingleFieldPivotAxisOutputItems(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	axis: 'row' | 'column',
	missingWarning: string,
): { ok: true; value: readonly PivotAxisOutputItem[] } | { ok: false; warning: string } {
	const fields = axis === 'row' ? pivot.rowFields : pivot.columnFields
	const realFields = fields.filter((field) => field.index >= 0)
	const axisFieldIndex = realFields.length === 1 ? realFields[0]?.index : undefined
	if (axisFieldIndex === undefined) return { ok: false, warning: missingWarning }
	const visibleValues = visiblePivotAxisFieldValues(cache, pivot, axisFieldIndex)
	if (!visibleValues.ok) return visibleValues
	const values =
		visibleValues.value.values.size > 0
			? visibleValues.value.values
			: new Set(
					(cache.fields[axisFieldIndex]?.sharedItems ?? [])
						.map((item) => item.value)
						.filter((value): value is string => value !== undefined),
				)
	const filterFieldIndex = visibleValues.value.fieldIndex ?? axisFieldIndex
	const output: PivotAxisOutputItem[] = []
	for (const value of values) {
		output.push({
			key: value,
			filters: new Map([[filterFieldIndex, new Set([value])]]),
		})
	}
	const hasGrandTotal =
		axis === 'row'
			? pivot.options?.rowGrandTotals !== false
			: pivot.options?.colGrandTotals !== false
	if (hasGrandTotal) {
		output.push({
			key: 'Grand Total',
			filters: values.size > 0 ? new Map([[filterFieldIndex, new Set(values)]]) : new Map(),
		})
	}
	return output.length > 0 ? { ok: true, value: output } : { ok: false, warning: missingWarning }
}

function buildTwoFieldSubtotalAxisOutputItems(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	axis: 'row' | 'column',
	missingWarning: string,
	crossAxisItems: readonly PivotAxisOutputItem[] | undefined,
): { ok: true; value: readonly PivotAxisOutputItem[] } | { ok: false; warning: string } {
	const fields = axis === 'row' ? pivot.rowFields : pivot.columnFields
	if (
		fields.length !== 2 ||
		!fields.every((field) => field.index >= 0) ||
		pivot.dataFields.length !== 1
	) {
		return { ok: false, warning: missingWarning }
	}
	const parentField = fields[0]
	const childField = fields[1]
	if (!parentField || !childField) return { ok: false, warning: missingWarning }
	const parentItemIndexes = visibleNonMissingPivotFieldItemIndexes(cache, pivot, parentField.index)
	const childItemIndexes = visibleNonMissingPivotFieldItemIndexes(cache, pivot, childField.index)
	if (parentItemIndexes.length === 0 || childItemIndexes.length === 0) {
		return { ok: false, warning: missingWarning }
	}
	const cacheRows = buildPivotCacheRows([cache], {})
	const output: PivotAxisOutputItem[] = []
	const seenLabels = new Map<string, number>()
	for (const parentItemIndex of parentItemIndexes) {
		const parent = resolvePivotAxisFieldItemFilter(cache, pivot, parentField.index, parentItemIndex)
		if (!parent.ok) return parent
		if (
			!cacheRows.some((row) =>
				pivotCacheRowMatchesAxisAndCrossAxis(row, parent.value.filters, crossAxisItems),
			)
		)
			continue
		for (const childItemIndex of childItemIndexes) {
			const child = resolvePivotAxisFieldItemFilter(cache, pivot, childField.index, childItemIndex)
			if (!child.ok) return child
			const filters = copyPivotAuditFilters(parent.value.filters)
			addPivotAxisFilters(filters, child.value.filters)
			if (
				!cacheRows.some((row) => pivotCacheRowMatchesAxisAndCrossAxis(row, filters, crossAxisItems))
			)
				continue
			output.push({
				key: uniquePivotAxisKey(child.value.label, seenLabels),
				filters,
			})
		}
		if (pivot.fields[parentField.index]?.defaultSubtotal !== false) {
			output.push({
				key: uniquePivotAxisKey(`${parent.value.label} Total`, seenLabels),
				filters: parent.value.filters,
			})
		}
	}
	const hasGrandTotal =
		axis === 'row'
			? pivot.options?.rowGrandTotals !== false
			: pivot.options?.colGrandTotals !== false
	if (hasGrandTotal) {
		const grand = buildPivotAxisGrandFilters(cache, pivot, fields)
		if (!grand.ok) return grand
		output.push({ key: uniquePivotAxisKey('Grand Total', seenLabels), filters: grand.value })
	}
	return output.length > 0 ? { ok: true, value: output } : { ok: false, warning: missingWarning }
}

function pivotCacheRowMatchesAxisAndCrossAxis(
	row: PivotCacheMaterializedRowInfo,
	filters: PivotAuditFilters,
	crossAxisItems: readonly PivotAxisOutputItem[] | undefined,
): boolean {
	if (!pivotCacheRowMatchesFilters(row, filters)) return false
	return (
		crossAxisItems === undefined ||
		crossAxisItems.some((item) => pivotCacheRowMatchesFilters(row, item.filters))
	)
}

function buildCompactTwoFieldRowAxisOutputItems(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	missingWarning: string,
): { ok: true; value: readonly PivotAxisOutputItem[] } | { ok: false; warning: string } {
	if (
		pivot.rowFields.length !== 2 ||
		!pivot.rowFields.every((field) => field.index >= 0) ||
		pivot.columnFields.length !== 0 ||
		pivot.dataFields.length !== 1 ||
		pivot.options?.compact !== true ||
		pivot.options?.outline !== true
	) {
		return { ok: false, warning: missingWarning }
	}
	const parentField = pivot.rowFields[0]
	const childField = pivot.rowFields[1]
	if (!parentField || !childField) return { ok: false, warning: missingWarning }
	const parentItemIndexes = visiblePivotFieldItemIndexes(pivot, parentField.index)
	const childItemIndexes = visiblePivotFieldItemIndexes(pivot, childField.index)
	if (parentItemIndexes.length === 0 || childItemIndexes.length === 0) {
		return { ok: false, warning: missingWarning }
	}
	const rows = buildPivotCacheRows([cache], {})
	const output: PivotAxisOutputItem[] = []
	const seenLabels = new Map<string, number>()
	for (const parentItemIndex of parentItemIndexes) {
		const parent = resolvePivotAxisFieldItemFilter(cache, pivot, parentField.index, parentItemIndex)
		if (!parent.ok) return parent
		if (!rows.some((row) => pivotCacheRowMatchesFilters(row, parent.value.filters))) continue
		output.push({
			key: uniquePivotAxisKey(parent.value.label, seenLabels),
			filters: parent.value.filters,
		})
		for (const childItemIndex of childItemIndexes) {
			const child = resolvePivotAxisFieldItemFilter(cache, pivot, childField.index, childItemIndex)
			if (!child.ok) return child
			const filters = copyPivotAuditFilters(parent.value.filters)
			addPivotAxisFilters(filters, child.value.filters)
			if (!rows.some((row) => pivotCacheRowMatchesFilters(row, filters))) continue
			output.push({
				key: uniquePivotAxisKey(child.value.label, seenLabels),
				filters,
			})
		}
	}
	if (pivot.options?.rowGrandTotals !== false) {
		const grand = buildPivotAxisGrandFilters(cache, pivot, pivot.rowFields)
		if (!grand.ok) return grand
		output.push({ key: uniquePivotAxisKey('Grand Total', seenLabels), filters: grand.value })
	}
	return output.length > 0 ? { ok: true, value: output } : { ok: false, warning: missingWarning }
}

function copyPivotAuditFilters(filters: PivotAuditFilters): Map<number, Set<string>> {
	return new Map(Array.from(filters, ([fieldIndex, values]) => [fieldIndex, new Set(values)]))
}

function visiblePivotFieldItemIndexes(pivot: PivotTableInfo, fieldIndex: number): number[] {
	return (pivot.fields[fieldIndex]?.items ?? [])
		.filter((item) => !item.hidden && !item.missing && item.cacheIndex !== undefined)
		.map((item) => item.index)
}

function visibleNonMissingPivotFieldItemIndexes(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	fieldIndex: number,
): number[] {
	return (pivot.fields[fieldIndex]?.items ?? [])
		.filter((item) => {
			if (item.hidden || item.missing || item.cacheIndex === undefined) return false
			return (
				pivotCacheSharedItemValue(cache, fieldIndex, item.cacheIndex) !==
				PIVOT_MISSING_ITEM_FILTER_VALUE
			)
		})
		.map((item) => item.index)
}

function buildPivotAxisGrandFilters(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	fields: readonly { readonly index: number }[],
): { ok: true; value: PivotAuditFilters } | { ok: false; warning: string } {
	const filters = new Map<number, Set<string>>()
	for (const axisField of fields) {
		if (axisField.index < 0) return { ok: false, warning: 'Data-field axis items are not audited.' }
		const values = visiblePivotAxisFieldValues(cache, pivot, axisField.index)
		if (!values.ok) return values
		if (values.value.fieldIndex !== undefined && values.value.values.size > 0) {
			filters.set(values.value.fieldIndex, values.value.values)
		}
	}
	return { ok: true, value: filters }
}

function visiblePivotAxisFieldValues(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	fieldIndex: number,
):
	| {
			ok: true
			value: {
				fieldIndex?: number
				values: Set<string>
			}
	  }
	| { ok: false; warning: string } {
	const pivotField = pivot.fields[fieldIndex]
	const items = pivotField?.items?.filter((item) => item.cacheIndex !== undefined) ?? []
	if (items.length === 0) return { ok: true, value: { values: new Set() } }
	const cacheField = cache.fields[fieldIndex]
	const group = cacheField?.fieldGroup
	if (group?.base !== undefined) {
		const values = new Set<string>()
		for (const item of items) {
			if (item.hidden || item.missing || item.cacheIndex === undefined) continue
			const groupedValues = pivotGroupedAxisBaseValues(cache, group, item.cacheIndex)
			if (!groupedValues.ok) return groupedValues
			for (const value of groupedValues.value) values.add(value)
		}
		if (values.size === 0) {
			return { ok: false, warning: 'Pivot grouped axis grand-total values were not resolved.' }
		}
		return { ok: true, value: { fieldIndex: group.base, values } }
	}
	const values = new Set<string>()
	for (const item of items) {
		if (item.hidden || item.missing || item.cacheIndex === undefined) continue
		const sharedItem = cache.fields[fieldIndex]?.sharedItems?.find(
			(entry) => entry.index === item.cacheIndex,
		)
		if (sharedItem?.value === undefined) {
			if (sharedItem?.kind === 'missing') {
				values.add(PIVOT_MISSING_ITEM_FILTER_VALUE)
				continue
			}
			return { ok: false, warning: 'Pivot axis grand-total shared item value was not resolved.' }
		}
		values.add(sharedItem.value)
	}
	return { ok: true, value: { fieldIndex, values } }
}

function resolvePivotAxisFieldItemFilter(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	fieldIndex: number,
	itemIndex: number,
):
	| {
			ok: true
			value: {
				label: string
				filters: PivotAuditFilters
			}
	  }
	| { ok: false; warning: string } {
	if (fieldIndex < 0) return { ok: false, warning: 'Data-field axis items are not audited.' }
	const pivotField = pivot.fields[fieldIndex]
	const pivotItem = pivotField?.items?.[itemIndex]
	if (!pivotItem || pivotItem.hidden || pivotItem.missing || pivotItem.cacheIndex === undefined) {
		return { ok: false, warning: 'Pivot axis field item cache value was not resolved.' }
	}
	const cacheField = cache.fields[fieldIndex]
	const group = cacheField?.fieldGroup
	if (group?.base !== undefined) {
		const label =
			group.groupItems?.find((entry) => entry.index === pivotItem.cacheIndex)?.value ??
			pivotItem.caption
		if (label === undefined) {
			return { ok: false, warning: 'Pivot grouped axis item label was not resolved.' }
		}
		const baseValues = pivotGroupedAxisBaseValues(cache, group, pivotItem.cacheIndex)
		if (!baseValues.ok) return baseValues
		if (baseValues.value.size === 0) {
			return { ok: false, warning: 'Pivot grouped axis item base values were not resolved.' }
		}
		return {
			ok: true,
			value: { label, filters: new Map([[group.base, baseValues.value]]) },
		}
	}
	const value = pivotCacheSharedItemValue(cache, fieldIndex, pivotItem.cacheIndex)
	if (value === undefined) {
		return { ok: false, warning: 'Pivot axis shared item value was not resolved.' }
	}
	return {
		ok: true,
		value: {
			label:
				pivotItem.caption ?? pivotCacheSharedItemLabel(cache, fieldIndex, pivotItem.cacheIndex),
			filters: new Map([[fieldIndex, new Set([value])]]),
		},
	}
}

function pivotGroupedAxisBaseValues(
	cache: PivotCacheInfo,
	group: NonNullable<PivotCacheInfo['fields'][number]['fieldGroup']>,
	groupItemIndex: number,
): { ok: true; value: Set<string> } | { ok: false; warning: string } {
	if (group.base === undefined) {
		return { ok: false, warning: 'Pivot grouped axis base field was not resolved.' }
	}
	const values = new Set<string>()
	for (const discreteItem of group.discreteItems ?? []) {
		if (discreteItem.value !== groupItemIndex) continue
		const value = pivotCacheSharedItemValue(cache, group.base, discreteItem.index)
		if (value !== undefined) values.add(value)
	}
	if (values.size > 0) return { ok: true, value: values }
	if (group.range?.groupBy === 'months') {
		for (const sharedItem of cache.fields[group.base]?.sharedItems ?? []) {
			if (sharedItem.kind !== 'date' || sharedItem.value === undefined) continue
			const month = pivotDateMonthIndex(sharedItem.value)
			if (month === groupItemIndex) values.add(sharedItem.value)
		}
		return { ok: true, value: values }
	}
	return { ok: true, value: values }
}

function pivotDateMonthIndex(value: string): number | undefined {
	const month = /^-?\d{4,}-(\d{2})-/u.exec(value)?.[1]
	if (!month) return undefined
	const parsed = Number(month)
	return Number.isInteger(parsed) && parsed >= 1 && parsed <= 12 ? parsed : undefined
}

function addPivotAxisFilters(target: Map<number, Set<string>>, source: PivotAuditFilters): void {
	for (const [fieldIndex, values] of source) {
		const existing = target.get(fieldIndex)
		if (!existing) {
			target.set(fieldIndex, new Set(values))
			continue
		}
		for (const value of Array.from(existing)) {
			if (!values.has(value)) existing.delete(value)
		}
	}
}

function uniquePivotAxisKey(label: string, seenLabels: Map<string, number>): string {
	const count = (seenLabels.get(label) ?? 0) + 1
	seenLabels.set(label, count)
	return count === 1 ? label : `${label} (${count})`
}

function aggregateSimplePivotOutput(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	rowFieldIndex: number,
	cacheRows: readonly PivotCacheMaterializedRowInfo[],
):
	| {
			ok: true
			value: {
				rowKeys: readonly string[]
				dataFieldNames: readonly string[]
				checkedValueCount: number
				values: ReadonlyMap<string, ReadonlyMap<string, number>>
			}
	  }
	| { ok: false; warning: string } {
	const pageFilters = buildSimplePivotPageFilters(cache, pivot)
	if (!pageFilters.ok) return pageFilters
	const dataFieldNames = pivot.dataFields.map((field, index) =>
		pivotDataFieldAuditName(cache, field, index),
	)
	const rows = buildPivotSimpleRowOutputItems(cache, pivot, rowFieldIndex)
	if (!rows.ok) return rows
	const output = new Map<string, Map<string, number>>()
	const baseTotals = new Map<string, Map<string, number>>()
	const averageTotals = new Map<string, Map<string, { sum: number; count: number }>>()
	const rowMatcher = buildPivotAxisOutputMatcher(rows.value)
	for (const row of cacheRows) {
		if (!pivotCacheRowMatchesFilters(row, pageFilters.value)) continue
		const rowItems = matchingPivotAxisOutputItems(row, rowMatcher)
		if (rowItems.length === 0) continue
		for (const rowItem of rowItems) addPivotBaseTotals(cache, row, baseTotals, rowItem.key)
		for (let i = 0; i < pivot.dataFields.length; i++) {
			const dataField = pivot.dataFields[i]
			const dataFieldName = dataFieldNames[i] ?? `DataField${i + 1}`
			if (!dataField) continue
			const field = cache.fields[dataField.fieldIndex]
			if (field?.formula) continue
			if (dataField.subtotal === 'average') {
				const value = numericPivotRowValue(row, dataField.fieldIndex)
				if (value === null) continue
				for (const rowItem of rowItems) {
					addPivotAverageSample(averageTotals, rowItem.key, dataFieldName, value)
				}
				continue
			}
			const measured = measurePivotDataField(cache, row, dataField)
			if (!measured.ok) return measured
			for (const rowItem of rowItems) {
				addPivotOutput(output, rowItem.key, dataFieldName, measured.value)
			}
		}
	}
	finalizePivotAverageOutput(output, averageTotals)
	for (let i = 0; i < pivot.dataFields.length; i++) {
		const dataField = pivot.dataFields[i]
		const dataFieldName = dataFieldNames[i] ?? `DataField${i + 1}`
		if (!dataField) continue
		const field = cache.fields[dataField.fieldIndex]
		if (!field?.formula) continue
		if (dataField.subtotal !== undefined && dataField.subtotal !== 'sum') {
			return {
				ok: false,
				warning: `Calculated pivot subtotal "${dataField.subtotal}" is not audited.`,
			}
		}
		for (const [rowLabel, valuesByName] of baseTotals) {
			const measured = measureCalculatedPivotField(valuesByName, field.formula)
			if (!measured.ok) return measured
			setPivotOutput(output, rowLabel, dataFieldName, measured.value)
		}
	}
	return {
		ok: true,
		value: {
			rowKeys: rows.value.map((row) => row.key),
			dataFieldNames,
			checkedValueCount: rows.value.length * dataFieldNames.length,
			values: output,
		},
	}
}

function buildSimplePivotOutputMaterializeUpdates(
	sheet: Workbook['sheets'][number],
	bounds: RangeRef,
	pivot: PivotTableInfo,
	expected: {
		readonly rowKeys: readonly string[]
		readonly dataFieldNames: readonly string[]
		readonly values: ReadonlyMap<string, ReadonlyMap<string, number>>
	},
	mode: PivotOutputMaterializeMode,
): { ok: true; value: readonly CellUpdate[] } | { ok: false; warning: string } {
	const headerRowOffset = Math.max(0, (pivot.location?.firstDataRow ?? 1) - 1)
	const dataStartRow = bounds.start.row + (pivot.location?.firstDataRow ?? headerRowOffset + 1)
	const dataStartCol = bounds.start.col + (pivot.location?.firstDataCol ?? 1)
	const headerRow = bounds.start.row + headerRowOffset
	const labelCol = dataStartCol > bounds.start.col ? dataStartCol - 1 : bounds.start.col
	if (
		headerRow < bounds.start.row ||
		headerRow > bounds.end.row ||
		dataStartRow > bounds.end.row ||
		dataStartCol > bounds.end.col ||
		labelCol < bounds.start.col ||
		labelCol > bounds.end.col ||
		dataStartRow + expected.rowKeys.length - 1 > bounds.end.row ||
		dataStartCol + expected.dataFieldNames.length - 1 > bounds.end.col
	) {
		return { ok: false, warning: 'Simple PivotTable output bounds were not resolved.' }
	}
	const updates: CellUpdate[] = []
	pushPivotMaterializeUpdate(
		sheet,
		updates,
		headerRow,
		labelCol,
		pivot.options?.rowHeaderCaption ?? 'Row Labels',
		mode,
	)
	for (let fieldIndex = 0; fieldIndex < expected.dataFieldNames.length; fieldIndex++) {
		pushPivotMaterializeUpdate(
			sheet,
			updates,
			headerRow,
			dataStartCol + fieldIndex,
			expected.dataFieldNames[fieldIndex] ?? `DataField${fieldIndex + 1}`,
			mode,
		)
	}
	for (let rowIndex = 0; rowIndex < expected.rowKeys.length; rowIndex++) {
		const rowKey = expected.rowKeys[rowIndex] ?? `Row ${rowIndex + 1}`
		const row = dataStartRow + rowIndex
		pushPivotMaterializeUpdate(sheet, updates, row, labelCol, rowKey, mode)
		for (let fieldIndex = 0; fieldIndex < expected.dataFieldNames.length; fieldIndex++) {
			const fieldName = expected.dataFieldNames[fieldIndex] ?? `DataField${fieldIndex + 1}`
			const value = expected.values.get(rowKey)?.get(fieldName) ?? 0
			pushPivotMaterializeUpdate(sheet, updates, row, dataStartCol + fieldIndex, value, mode)
		}
	}
	return { ok: true, value: updates }
}

function pushPivotMaterializeUpdate(
	sheet: Workbook['sheets'][number],
	updates: CellUpdate[],
	row: number,
	col: number,
	value: string | number | null,
	mode: PivotOutputMaterializeMode,
): void {
	const existing = sheet.cells.get(row, col)?.value ?? EMPTY
	if (mode === 'missing' && existing.kind !== 'empty') return
	if (mode === 'mismatches' && pivotMaterializedCellMatches(existing, value)) return
	updates.push({ ref: `${indexToColumn(col)}${row + 1}`, value })
}

function pivotMaterializedCellMatches(existing: CellValue, value: string | number | null): boolean {
	if (value === null) return existing.kind === 'empty'
	if (typeof value === 'number') return numericCellMatches(existing, value)
	return cellText(existing) === value
}

type PivotAuditFilters = ReadonlyMap<number, ReadonlySet<string>>

type PivotAxisOutputMatcher<T extends PivotAxisOutputItem> = {
	readonly byFieldValue: ReadonlyMap<number, ReadonlyMap<string, readonly T[]>>
	readonly fallback: readonly T[]
}

function buildPivotAxisOutputMatcher<T extends PivotAxisOutputItem>(
	items: readonly T[],
): PivotAxisOutputMatcher<T> {
	const byFieldValue = new Map<number, Map<string, T[]>>()
	const fallback: T[] = []
	for (const item of items) {
		if (item.filters.size !== 1) {
			fallback.push(item)
			continue
		}
		const [filter] = item.filters
		if (!filter) {
			fallback.push(item)
			continue
		}
		const [fieldIndex, values] = filter
		for (const value of values) {
			let byValue = byFieldValue.get(fieldIndex)
			if (!byValue) {
				byValue = new Map()
				byFieldValue.set(fieldIndex, byValue)
			}
			const outputItems = byValue.get(value)
			if (outputItems) {
				outputItems.push(item)
				continue
			}
			byValue.set(value, [item])
		}
	}
	return { byFieldValue, fallback }
}

function matchingPivotAxisOutputItems<T extends PivotAxisOutputItem>(
	row: PivotCacheMaterializedRowInfo,
	matcher: PivotAxisOutputMatcher<T>,
): readonly T[] {
	const output: T[] = []
	for (const [fieldIndex, byValue] of matcher.byFieldValue) {
		const actualValue = row.values.find((value) => value.fieldIndex === fieldIndex)
		const actual = actualValue ? pivotCacheDecodedFilterValue(actualValue) : undefined
		if (actual === undefined) continue
		const matched = byValue.get(actual)
		if (matched) output.push(...matched)
	}
	for (const item of matcher.fallback) {
		if (pivotCacheRowMatchesFilters(row, item.filters)) output.push(item)
	}
	return output
}

function buildPivotSimpleRowOutputItems(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	rowFieldIndex: number,
): { ok: true; value: readonly PivotAxisOutputItem[] } | { ok: false; warning: string } {
	const rows = buildPivotAxisOutputItems(cache, pivot, 'row')
	if (rows.ok || rows.warning !== 'Pivot row output items were not found.') return rows
	const visibleValues = visiblePivotAxisFieldValues(cache, pivot, rowFieldIndex)
	if (!visibleValues.ok) return visibleValues
	const values =
		visibleValues.value.values.size > 0
			? visibleValues.value.values
			: new Set(
					(cache.fields[rowFieldIndex]?.sharedItems ?? [])
						.map((item) => item.value)
						.filter((value): value is string => value !== undefined),
				)
	const output: PivotAxisOutputItem[] = []
	for (const value of values) {
		output.push({
			key: value,
			filters: new Map([[visibleValues.value.fieldIndex ?? rowFieldIndex, new Set([value])]]),
		})
	}
	if (pivot.options?.rowGrandTotals !== false) {
		output.push({
			key: 'Grand Total',
			filters:
				values.size > 0
					? new Map([[visibleValues.value.fieldIndex ?? rowFieldIndex, new Set(values)]])
					: new Map(),
		})
	}
	return output.length > 0
		? { ok: true, value: output }
		: { ok: false, warning: 'Pivot row output items were not found.' }
}

function buildSimplePivotPageFilters(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
): { ok: true; value: PivotAuditFilters } | { ok: false; warning: string } {
	const filters = new Map<number, Set<string>>()
	for (const pageField of pivot.pageFields) {
		if (pageField.index < 0) {
			return { ok: false, warning: 'Data-field page filters are not audited.' }
		}
		const field = pivot.fields[pageField.index]
		const allowed = pivotPageFieldAllowedValues(cache, pageField.index, field, pageField.item)
		if (!allowed.ok) return allowed
		if (allowed.value.values.size > 0) {
			filters.set(allowed.value.fieldIndex ?? pageField.index, allowed.value.values)
		}
	}
	return { ok: true, value: filters }
}

function pivotPageFieldAllowedValues(
	cache: PivotCacheInfo,
	fieldIndex: number,
	field: PivotFieldInfo | undefined,
	selectedItemIndex: number | undefined,
):
	| {
			ok: true
			value: {
				fieldIndex?: number
				values: Set<string>
			}
	  }
	| { ok: false; warning: string } {
	const cacheField = cache.fields[fieldIndex]
	const group = cacheField?.fieldGroup
	if (selectedItemIndex !== undefined) {
		const selected = field?.items?.[selectedItemIndex]
		if (!selected || selected.hidden || selected.missing || selected.cacheIndex === undefined) {
			return { ok: false, warning: 'Pivot page filter selected item was not resolved.' }
		}
		if (group?.base !== undefined) {
			const values = pivotGroupedAxisBaseValues(cache, group, selected.cacheIndex)
			if (!values.ok) return values
			return { ok: true, value: { fieldIndex: group.base, values: values.value } }
		}
		const value = pivotCacheSharedItemValue(cache, fieldIndex, selected.cacheIndex)
		if (value === undefined) {
			return { ok: false, warning: 'Pivot page filter selected cache item was not resolved.' }
		}
		return { ok: true, value: { fieldIndex, values: new Set([value]) } }
	}
	const items = field?.items?.filter((item) => item.cacheIndex !== undefined) ?? []
	if (items.length === 0 || items.every((item) => !item.hidden && !item.missing)) {
		return { ok: true, value: { values: new Set() } }
	}
	const visibleValues = new Set<string>()
	for (const item of items) {
		if (item.hidden || item.missing || item.cacheIndex === undefined) continue
		if (group?.base !== undefined) {
			const values = pivotGroupedAxisBaseValues(cache, group, item.cacheIndex)
			if (!values.ok) return values
			for (const value of values.value) visibleValues.add(value)
			continue
		}
		const value = pivotCacheSharedItemValue(cache, fieldIndex, item.cacheIndex)
		if (value === undefined) {
			return { ok: false, warning: 'Pivot page filter visible cache item was not resolved.' }
		}
		visibleValues.add(value)
	}
	if (visibleValues.size === items.length && group?.base === undefined) {
		return { ok: true, value: { values: new Set() } }
	}
	if (
		visibleValues.size === 1 &&
		visibleValues.has(PIVOT_MISSING_ITEM_FILTER_VALUE) &&
		pageFieldVisibleMissingItemIsTrailingPlaceholder(cacheField, items)
	) {
		return { ok: true, value: { values: new Set() } }
	}
	return { ok: true, value: { fieldIndex: group?.base ?? fieldIndex, values: visibleValues } }
}

function pageFieldVisibleMissingItemIsTrailingPlaceholder(
	cacheField: PivotCacheInfo['fields'][number] | undefined,
	items: readonly PivotFieldItemInfo[],
): boolean {
	const visibleItems = items.filter(
		(item) => !item.hidden && !item.missing && item.cacheIndex !== undefined,
	)
	if (visibleItems.length !== 1) return false
	const visibleItem = visibleItems[0]
	if (!visibleItem || visibleItem.cacheIndex === undefined) return false
	const sharedItems = cacheField?.sharedItems ?? []
	const sharedItem = sharedItems.find((entry) => entry.index === visibleItem.cacheIndex)
	if (sharedItem?.kind !== 'missing') return false
	const maxSharedIndex = Math.max(...sharedItems.map((entry) => entry.index))
	return Number.isFinite(maxSharedIndex) && visibleItem.cacheIndex === maxSharedIndex
}

function pivotCacheSharedItemValue(
	cache: PivotCacheInfo,
	fieldIndex: number,
	sharedItemIndex: number,
): string | undefined {
	const sharedItem = cache.fields[fieldIndex]?.sharedItems?.find(
		(entry) => entry.index === sharedItemIndex,
	)
	if (sharedItem?.value !== undefined) return sharedItem.value
	if (sharedItem?.kind === 'missing') return PIVOT_MISSING_ITEM_FILTER_VALUE
	return undefined
}

function pivotCacheSharedItemLabel(
	cache: PivotCacheInfo,
	fieldIndex: number,
	sharedItemIndex: number,
): string {
	const sharedItem = cache.fields[fieldIndex]?.sharedItems?.find(
		(entry) => entry.index === sharedItemIndex,
	)
	if (sharedItem?.kind === 'missing') return '(blank)'
	if (sharedItem?.kind === 'boolean') return sharedItem.value === '0' ? 'FALSE' : 'TRUE'
	if (sharedItem?.kind === 'number' && sharedItem.value !== undefined) {
		const numeric = Number(sharedItem.value)
		if (Number.isFinite(numeric)) return String(numeric)
	}
	return sharedItem?.value ?? ''
}

function pivotCacheRowMatchesFilters(
	row: PivotCacheMaterializedRowInfo,
	filters: PivotAuditFilters,
): boolean {
	for (const [fieldIndex, expectedValues] of filters) {
		const actualValue = row.values.find((value) => value.fieldIndex === fieldIndex)
		const actual = actualValue ? pivotCacheDecodedFilterValue(actualValue) : undefined
		if (actual === undefined || !expectedValues.has(actual)) return false
	}
	return true
}

function pivotCacheDecodedFilterValue(value: PivotCacheDecodedValueInfo): string | undefined {
	if (value.value !== undefined) return value.value
	return value.kind === 'missing' || value.sharedItemKind === 'missing'
		? PIVOT_MISSING_ITEM_FILTER_VALUE
		: undefined
}

function measurePivotDataField(
	cache: PivotCacheInfo,
	row: PivotCacheMaterializedRowInfo,
	dataField: PivotTableInfo['dataFields'][number],
): { ok: true; value: number } | { ok: false; warning: string } {
	const field = cache.fields[dataField.fieldIndex]
	if (!field)
		return { ok: false, warning: `Pivot data field ${dataField.fieldIndex} was not found.` }
	if (dataField.subtotal === 'count') {
		return {
			ok: true,
			value: pivotCountFieldHasValue(row, field.index, field.fieldGroup?.base) ? 1 : 0,
		}
	}
	if (dataField.subtotal !== undefined && dataField.subtotal !== 'sum') {
		return { ok: false, warning: `Pivot subtotal "${dataField.subtotal}" is not audited.` }
	}
	if (field.formula)
		return { ok: false, warning: 'Calculated pivot fields require aggregate audit.' }
	const value = numericPivotRowValue(row, field.index)
	if (value === null) return { ok: true, value: 0 }
	return { ok: true, value }
}

function pivotCountFieldHasValue(
	row: PivotCacheMaterializedRowInfo,
	fieldIndex: number,
	groupBaseFieldIndex: number | undefined,
): boolean {
	if (
		row.values.some(
			(value) => value.fieldIndex === fieldIndex && pivotCacheDecodedHasCountValue(value),
		)
	) {
		return true
	}
	return (
		groupBaseFieldIndex !== undefined &&
		row.values.some(
			(value) => value.fieldIndex === groupBaseFieldIndex && pivotCacheDecodedHasCountValue(value),
		)
	)
}

function pivotCacheDecodedHasCountValue(value: PivotCacheDecodedValueInfo): boolean {
	return value.value !== undefined
}

function measureCalculatedPivotField(
	valuesByName: ReadonlyMap<string, number>,
	formula: string,
): { ok: true; value: number } | { ok: false; warning: string } {
	const parsed = parseFormula(formula)
	if (!parsed.ok)
		return { ok: false, warning: `Pivot calculated field formula did not parse: ${formula}` }
	const measured = evalSimplePivotFormula(parsed.value, valuesByName)
	if (measured === null) {
		return {
			ok: false,
			warning: `Pivot calculated field formula is not audit-supported: ${formula}`,
		}
	}
	return { ok: true, value: measured }
}

function addPivotBaseTotals(
	cache: PivotCacheInfo,
	row: PivotCacheMaterializedRowInfo,
	output: Map<string, Map<string, number>>,
	rowLabel: string,
): void {
	let byField = output.get(rowLabel)
	if (!byField) {
		byField = new Map()
		output.set(rowLabel, byField)
	}
	for (const value of row.values) {
		const field = cache.fields[value.fieldIndex]
		if (!field?.name) continue
		const numeric = numericPivotRowValue(row, value.fieldIndex)
		if (numeric !== null) {
			const key = normalizePivotAuditText(field.name)
			byField.set(key, (byField.get(key) ?? 0) + numeric)
		}
	}
}

function numericPivotRowValue(
	row: PivotCacheMaterializedRowInfo,
	fieldIndex: number,
): number | null {
	const value = row.values.find((entry) => entry.fieldIndex === fieldIndex)
	if (!value?.value || value.kind !== 'number') return null
	const numeric = Number(value.value)
	return Number.isFinite(numeric) ? numeric : null
}

function evalSimplePivotFormula(
	node: FormulaNode,
	valuesByName: ReadonlyMap<string, number>,
): number | null {
	switch (node.type) {
		case 'number':
			return node.value
		case 'name':
			return valuesByName.get(normalizePivotAuditText(node.name)) ?? null
		case 'binary': {
			const left = evalSimplePivotFormula(node.left, valuesByName)
			const right = evalSimplePivotFormula(node.right, valuesByName)
			if (left === null || right === null) return null
			switch (node.op) {
				case '+':
					return left + right
				case '-':
					return left - right
				case '*':
					return left * right
				case '/':
					return right === 0 ? null : left / right
				default:
					return null
			}
		}
		default:
			return null
	}
}

function addPivotOutput(
	output: Map<string, Map<string, number>>,
	rowLabel: string,
	dataField: string,
	value: number,
): void {
	let byField = output.get(rowLabel)
	if (!byField) {
		byField = new Map()
		output.set(rowLabel, byField)
	}
	byField.set(dataField, (byField.get(dataField) ?? 0) + value)
}

function addPivotAverageSample(
	output: Map<string, Map<string, { sum: number; count: number }>>,
	rowLabel: string,
	dataField: string,
	value: number,
): void {
	let byField = output.get(rowLabel)
	if (!byField) {
		byField = new Map()
		output.set(rowLabel, byField)
	}
	const existing = byField.get(dataField) ?? { sum: 0, count: 0 }
	byField.set(dataField, { sum: existing.sum + value, count: existing.count + 1 })
}

function finalizePivotAverageOutput(
	output: Map<string, Map<string, number>>,
	averageTotals: ReadonlyMap<string, ReadonlyMap<string, { sum: number; count: number }>>,
): void {
	for (const [rowLabel, byField] of averageTotals) {
		for (const [dataField, aggregate] of byField) {
			if (aggregate.count > 0)
				setPivotOutput(output, rowLabel, dataField, aggregate.sum / aggregate.count)
		}
	}
}

function setPivotOutput(
	output: Map<string, Map<string, number>>,
	rowLabel: string,
	dataField: string,
	value: number,
): void {
	let byField = output.get(rowLabel)
	if (!byField) {
		byField = new Map()
		output.set(rowLabel, byField)
	}
	byField.set(dataField, value)
}

function readSimplePivotOutput(
	workbook: Workbook,
	sheetId: string,
	pivot: PivotTableInfo,
	dataFieldNames: readonly string[],
):
	| { ok: true; value: ReadonlyMap<string, ReadonlyMap<string, { ref: string; value: CellValue }>> }
	| { ok: false; warning: string } {
	if (!pivot.locationRef) return { ok: false, warning: 'Pivot table has no output location.' }
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return { ok: false, warning: `Pivot output range is invalid: ${pivot.locationRef}` }
	}
	const sheet = workbook.sheets.find((entry) => entry.id === sheetId)
	if (!sheet) return { ok: false, warning: 'Pivot output sheet was not loaded.' }
	const headerNamesByNormalized = pivotDataFieldHeaderNamesByNormalized(dataFieldNames)
	const headers = new Map<string, { row: number; col: number }>()
	for (let row = bounds.start.row; row <= bounds.end.row; row++) {
		for (let col = bounds.start.col; col <= bounds.end.col; col++) {
			const text = cellText(sheet.cells.get(row, col)?.value ?? EMPTY)
			const fieldName = headerNamesByNormalized.get(normalizePivotAuditText(text))
			if (fieldName) headers.set(fieldName, { row, col })
		}
	}
	if (headers.size !== dataFieldNames.length) {
		return { ok: false, warning: 'Pivot output data-field headers were not found.' }
	}
	const headerRow = Math.min(...Array.from(headers.values(), (entry) => entry.row))
	const dataStartCol = bounds.start.col + (pivot.location?.firstDataCol ?? 1)
	const output = new Map<string, Map<string, { ref: string; value: CellValue }>>()
	for (let row = headerRow + 1; row <= bounds.end.row; row++) {
		const labelValue = sheet.cells.get(row, bounds.start.col)?.value ?? EMPTY
		const rawLabel = cellText(labelValue)
		if (!rawLabel) continue
		const label = normalizeGrandTotalLabel(rawLabel)
		const byField = new Map<string, { ref: string; value: CellValue }>()
		for (const fieldName of dataFieldNames) {
			const header = headers.get(fieldName)
			if (!header) continue
			const valueCol = Math.max(header.col, dataStartCol)
			const value = sheet.cells.get(row, valueCol)?.value ?? EMPTY
			byField.set(fieldName, {
				ref: `${indexToColumn(valueCol)}${row + 1}`,
				value,
			})
		}
		output.set(label, byField)
	}
	return { ok: true, value: output }
}

function readDataFieldsOnRowsPivotOutput(
	workbook: Workbook,
	sheetId: string,
	pivot: PivotTableInfo,
	dataFieldNames: readonly string[],
	columnKeys: readonly string[],
):
	| { ok: true; value: ReadonlyMap<string, ReadonlyMap<string, { ref: string; value: CellValue }>> }
	| { ok: false; warning: string } {
	if (!pivot.locationRef) return { ok: false, warning: 'Pivot table has no output location.' }
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return { ok: false, warning: `Pivot output range is invalid: ${pivot.locationRef}` }
	}
	const sheet = workbook.sheets.find((entry) => entry.id === sheetId)
	if (!sheet) return { ok: false, warning: 'Pivot output sheet was not loaded.' }
	const dataStartRow = bounds.start.row + (pivot.location?.firstDataRow ?? 1)
	const dataStartCol = bounds.start.col + (pivot.location?.firstDataCol ?? 1)
	if (
		dataStartRow > bounds.end.row ||
		dataStartCol > bounds.end.col ||
		dataStartCol + columnKeys.length - 1 > bounds.end.col
	) {
		return { ok: false, warning: 'Pivot data-fields-on-rows output bounds were not resolved.' }
	}
	const namesByNormalized = new Map(
		dataFieldNames.map((name) => [normalizePivotAuditText(name), name] as const),
	)
	const output = new Map<string, Map<string, { ref: string; value: CellValue }>>()
	for (let row = dataStartRow; row <= bounds.end.row; row++) {
		const rowLabel = cellText(sheet.cells.get(row, bounds.start.col)?.value ?? EMPTY)
		const dataFieldName = namesByNormalized.get(normalizePivotAuditText(rowLabel))
		if (!dataFieldName) continue
		const byColumn = new Map<string, { ref: string; value: CellValue }>()
		for (let i = 0; i < columnKeys.length; i++) {
			const col = dataStartCol + i
			const columnKey = columnKeys[i] ?? `Column ${i + 1}`
			byColumn.set(columnKey, {
				ref: `${indexToColumn(col)}${row + 1}`,
				value: sheet.cells.get(row, col)?.value ?? EMPTY,
			})
		}
		output.set(dataFieldName, byColumn)
	}
	if (output.size !== dataFieldNames.length) {
		return { ok: false, warning: 'Pivot data-fields-on-rows labels were not found.' }
	}
	return { ok: true, value: output }
}

function readMultiRowAxisPivotOutput(
	workbook: Workbook,
	sheetId: string,
	pivot: PivotTableInfo,
	rowKeys: readonly string[],
	columnKeys: readonly string[],
):
	| { ok: true; value: ReadonlyMap<string, ReadonlyMap<string, { ref: string; value: CellValue }>> }
	| { ok: false; warning: string } {
	if (!pivot.locationRef) return { ok: false, warning: 'Pivot table has no output location.' }
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return { ok: false, warning: `Pivot output range is invalid: ${pivot.locationRef}` }
	}
	const sheet = workbook.sheets.find((entry) => entry.id === sheetId)
	if (!sheet) return { ok: false, warning: 'Pivot output sheet was not loaded.' }
	const dataStartRow = bounds.start.row + (pivot.location?.firstDataRow ?? 1)
	const dataStartCol = bounds.start.col + (pivot.location?.firstDataCol ?? 1)
	if (
		dataStartRow > bounds.end.row ||
		dataStartCol > bounds.end.col ||
		dataStartRow + rowKeys.length - 1 > bounds.end.row ||
		dataStartCol + columnKeys.length - 1 > bounds.end.col
	) {
		return { ok: false, warning: 'Multi-row pivot output bounds were not resolved.' }
	}
	const output = new Map<string, Map<string, { ref: string; value: CellValue }>>()
	for (let rowIndex = 0; rowIndex < rowKeys.length; rowIndex++) {
		const row = dataStartRow + rowIndex
		const rowKey = rowKeys[rowIndex] ?? `Row ${rowIndex + 1}`
		const byColumn = new Map<string, { ref: string; value: CellValue }>()
		for (let columnIndex = 0; columnIndex < columnKeys.length; columnIndex++) {
			const col = dataStartCol + columnIndex
			const columnKey = columnKeys[columnIndex] ?? `Column ${columnIndex + 1}`
			byColumn.set(columnKey, {
				ref: `${indexToColumn(col)}${row + 1}`,
				value: sheet.cells.get(row, col)?.value ?? EMPTY,
			})
		}
		output.set(rowKey, byColumn)
	}
	return { ok: true, value: output }
}

function pivotDataFieldHeaderNamesByNormalized(
	dataFieldNames: readonly string[],
): ReadonlyMap<string, string> {
	const aliases = new Map<string, string>()
	for (const fieldName of dataFieldNames) {
		addUniquePivotHeaderAlias(aliases, normalizePivotAuditText(fieldName), fieldName)
		const inferred = /^(?:sum|count|average) - (.+)$/iu.exec(fieldName)?.[1]
		if (inferred) addUniquePivotHeaderAlias(aliases, normalizePivotAuditText(inferred), fieldName)
	}
	for (const [alias, fieldName] of Array.from(aliases)) {
		if (fieldName === '') aliases.delete(alias)
	}
	return aliases
}

function addUniquePivotHeaderAlias(
	aliases: Map<string, string>,
	alias: string,
	fieldName: string,
): void {
	if (!alias) return
	const existing = aliases.get(alias)
	if (existing === undefined) {
		aliases.set(alias, fieldName)
		return
	}
	if (existing !== fieldName) aliases.set(alias, '')
}

function comparePivotOutput(
	expected: ReadonlyMap<string, ReadonlyMap<string, number>>,
	actual: ReadonlyMap<string, ReadonlyMap<string, { ref: string; value: CellValue }>>,
): PivotOutputAuditMismatchInfo[] {
	const mismatches: PivotOutputAuditMismatchInfo[] = []
	for (const [rowLabel, expectedFields] of expected) {
		const actualFields = actual.get(rowLabel)
		for (const [dataField, expectedValue] of expectedFields) {
			const actualCell = actualFields?.get(dataField)
			if (!actualCell || !numericCellMatches(actualCell.value, expectedValue)) {
				mismatches.push({
					...(actualCell ? { ref: actualCell.ref, actual: actualCell.value } : {}),
					rowLabel,
					dataField,
					expected: expectedValue,
				})
			}
		}
	}
	return mismatches
}

function numericCellMatches(value: CellValue, expected: number): boolean {
	if (value.kind === 'empty') return expected === 0
	if (value.kind !== 'number') return false
	const tolerance = 1e-12 * Math.max(1, Math.abs(value.value), Math.abs(expected))
	return Math.abs(value.value - expected) <= tolerance
}

function cellText(value: CellValue): string {
	switch (value.kind) {
		case 'string':
			return value.value
		case 'number':
			return String(value.value)
		case 'boolean':
			return value.value ? 'TRUE' : 'FALSE'
		case 'error':
			return value.value
		default:
			return ''
	}
}

function normalizeGrandTotalLabel(label: string): string {
	const normalized = normalizePivotAuditText(label)
	return normalized === 'grand total' ||
		normalized === 'total result' ||
		normalized === 'gesamtergebnis' ||
		normalized === 'общий итог'
		? 'Grand Total'
		: label
}

function normalizeSavedPivotLabel(label: string): string {
	return normalizePivotAuditText(normalizeGrandTotalLabel(label))
}

function normalizePivotAuditText(value: string): string {
	return value.trim().replace(/\s+/g, ' ').toLowerCase()
}

function copySlicerCacheInfo(cache: SlicerCacheInfo): SlicerCacheInfo {
	const items = cache.items?.map((item) => ({ ...item }))
	return {
		...cache,
		pivotTableNames: [...cache.pivotTableNames],
		...(items ? { items } : {}),
	}
}

function copyPivotTableInfo(pivot: PivotTableInfo): PivotTableInfo {
	return clonePivotTableInfo(pivot)
}

function buildGetPivotDataResult(query: GetPivotDataQuery, workbook: Workbook): GetPivotDataResult {
	const normalizedPivotTable = query.pivotTable?.toLowerCase()
	const filters = query.filters ?? []
	const matches = workbook.pivotTables.flatMap((pivot) => {
		if (normalizedPivotTable && pivot.name?.toLowerCase() !== normalizedPivotTable) return []
		const dataField = pivot.dataFields.find((field) => {
			const sourceFieldName = pivot.fields.find(
				(pivotField) => pivotField.index === field.fieldIndex,
			)?.name
			return (
				pivotDataFieldCaptionsMatch(query.dataField, field.name) ||
				pivotDataFieldCaptionsMatch(query.dataField, sourceFieldName)
			)
		})
		if (!dataField) return []
		const pivotFieldNames = new Set(
			pivot.fields.flatMap((field) => (field.name ? [field.name.toLowerCase()] : [])),
		)
		const matchedFilters = filters.filter((filter) =>
			pivotFieldNames.has(filter.field.toLowerCase()),
		)
		const unmatchedFilters = filters.filter(
			(filter) => !pivotFieldNames.has(filter.field.toLowerCase()),
		)
		const output =
			unmatchedFilters.length === 0
				? resolveSavedPivotOutput(workbook, pivot, query.dataField, filters)
				: undefined
		return [
			{
				pivotTable: copyPivotTableInfo(pivot),
				dataField: { ...dataField },
				matchedFilters,
				unmatchedFilters,
				...(output ? { output } : {}),
			},
		]
	})
	const canResolveOutput = matches.some((match) => match.output !== undefined)
	const warnings: string[] = []
	if (matches.length === 0) {
		warnings.push('No matching pivot table/data field metadata was found.')
	}
	if (matches.length > 0 && !canResolveOutput) {
		warnings.push(
			'GETPIVOTDATA metadata can be resolved, but no matching saved pivot output value was found.',
		)
	}
	if (matches.some((match) => match.unmatchedFilters.length > 0)) {
		warnings.push(
			'Some requested field/item filters are not present in inspectable pivot metadata.',
		)
	}
	warnings.push('GETPIVOTDATA saved output values are read from the workbook, not recalculated.')
	return {
		query,
		matches,
		canResolveOutput,
		warnings,
	}
}

function resolveSavedPivotOutput(
	workbook: Workbook,
	pivot: PivotTableInfo,
	dataField: string,
	filters: readonly GetPivotDataFilter[],
): { sheetName: string; ref: string; value: CellValue } | undefined {
	if (!pivot.locationRef) return undefined
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return undefined
	}
	const sheetIndex = workbook.sheets.findIndex((entry) => entry.name === pivot.sheetName)
	const sheet = workbook.sheets[sheetIndex]
	if (!sheet) return undefined
	const header = findSavedPivotDataHeader(sheet, bounds, dataField)
	if (!header) return undefined
	const cache = workbook.pivotCaches.find((entry) => entry.cacheId === pivot.cacheId)
	const visibleFilters = resolveSavedPivotPageFilters(pivot, cache, filters)
	if (!visibleFilters) return undefined
	const axisFilters = splitSavedPivotAxisFilters(pivot, cache, visibleFilters)
	const fieldColumns = axisFilters.row.map((filter) => ({
		filter,
		col: findSavedPivotRowFieldColumn(sheet, bounds, header, filter.field),
	}))
	if (fieldColumns.some((entry) => entry.col === undefined)) return undefined
	const columnHeaderEndRow =
		(findSavedPivotDataStartRow(sheet, bounds, pivot, cache, header) ?? bounds.end.row + 1) - 1
	for (let row = header.row + 1; row <= bounds.end.row; row++) {
		if (!savedPivotRowMatchesFilters(sheet, bounds, row, fieldColumns)) continue
		const col = findSavedPivotValueColumn(
			sheet,
			bounds,
			header,
			columnHeaderEndRow,
			axisFilters.column,
			row,
			fieldColumns.length,
		)
		if (col === undefined) return undefined
		return {
			sheetName: pivot.sheetName,
			ref: `${indexToColumn(col)}${row + 1}`,
			value: sheet.cells.get(row, col)?.value ?? EMPTY,
		}
	}
	return undefined
}

function resolveSavedPivotPageFilters(
	pivot: PivotTableInfo,
	cache: PivotCacheInfo | undefined,
	filters: readonly GetPivotDataFilter[],
): readonly GetPivotDataFilter[] | null {
	if (pivot.pageFields.length === 0 || filters.length === 0) return filters
	const visibleFilters: GetPivotDataFilter[] = []
	for (const filter of filters) {
		const pageField = pivot.pageFields.find((candidate) =>
			savedPivotFieldReferenceCaptionsMatch(pivot, cache, candidate, filter.field),
		)
		if (!pageField) {
			visibleFilters.push(filter)
			continue
		}
		const selected = selectedSavedPivotPageFieldItem(pivot, cache, pageField)
		if (!selected || normalizePivotAuditText(selected) !== normalizePivotAuditText(filter.item)) {
			return null
		}
	}
	return visibleFilters
}

function selectedSavedPivotPageFieldItem(
	pivot: PivotTableInfo,
	cache: PivotCacheInfo | undefined,
	pageField: PivotFieldReference,
): string | null {
	if (pageField.index < 0 || pageField.item === undefined) return null
	const pivotItem = pivot.fields[pageField.index]?.items?.[pageField.item]
	if (!pivotItem || pivotItem.hidden || pivotItem.missing || pivotItem.cacheIndex === undefined) {
		return null
	}
	if (!cache) return pivotItem.caption ?? null
	return pivotCacheSharedItemLabel(cache, pageField.index, pivotItem.cacheIndex)
}

function splitSavedPivotAxisFilters(
	pivot: PivotTableInfo,
	cache: PivotCacheInfo | undefined,
	filters: readonly GetPivotDataFilter[],
): { readonly row: readonly GetPivotDataFilter[]; readonly column: readonly GetPivotDataFilter[] } {
	const row: GetPivotDataFilter[] = []
	const column: GetPivotDataFilter[] = []
	for (const filter of filters) {
		if (savedPivotFieldReferenceListMatches(pivot, cache, pivot.columnFields, filter.field)) {
			column.push(filter)
		} else {
			row.push(filter)
		}
	}
	return { row, column }
}

function savedPivotFieldReferenceListMatches(
	pivot: PivotTableInfo,
	cache: PivotCacheInfo | undefined,
	references: readonly PivotFieldReference[],
	fieldName: string,
): boolean {
	return references.some((reference) =>
		savedPivotFieldReferenceCaptionsMatch(pivot, cache, reference, fieldName),
	)
}

function savedPivotFieldReferenceCaptionsMatch(
	pivot: PivotTableInfo,
	cache: PivotCacheInfo | undefined,
	reference: PivotFieldReference,
	fieldName: string,
): boolean {
	const field = pivot.fields[reference.index]
	const cacheField = cache?.fields[reference.index]
	const normalized = normalizePivotAuditText(fieldName)
	return [field?.name, cacheField?.name, reference.name, reference.caption]
		.filter((caption): caption is string => caption !== undefined)
		.some((caption) => normalizePivotAuditText(caption) === normalized)
}

function findSavedPivotDataHeader(
	sheet: Workbook['sheets'][number],
	bounds: RangeRef,
	dataField: string,
): { row: number; col: number } | undefined {
	const maxHeaderRow = Math.min(bounds.end.row, bounds.start.row + 8)
	for (let row = bounds.start.row; row <= maxHeaderRow; row++) {
		for (let col = bounds.start.col; col <= bounds.end.col; col++) {
			const header = cellText(sheet.cells.get(row, col)?.value ?? EMPTY)
			if (pivotDataFieldCaptionsMatch(dataField, header)) {
				return { row, col }
			}
		}
	}
	return undefined
}

function findSavedPivotFieldColumn(
	sheet: Workbook['sheets'][number],
	bounds: RangeRef,
	headerRow: number,
	dataCol: number,
	fieldName: string,
): number | undefined {
	const normalizedFieldName = normalizePivotAuditText(fieldName)
	for (let col = bounds.start.col; col < dataCol; col++) {
		const header = normalizePivotAuditText(
			cellText(sheet.cells.get(headerRow, col)?.value ?? EMPTY),
		)
		if (header === normalizedFieldName) return col
	}
	return dataCol > bounds.start.col ? bounds.start.col : undefined
}

function findSavedPivotRowFieldColumn(
	sheet: Workbook['sheets'][number],
	bounds: RangeRef,
	header: { row: number; col: number },
	fieldName: string,
): number | undefined {
	const normalizedFieldName = normalizePivotAuditText(fieldName)
	const maxHeaderRow = Math.min(bounds.end.row, header.row + 3)
	const maxLabelCol = header.col > bounds.start.col ? header.col - 1 : bounds.start.col
	for (let row = bounds.start.row; row <= maxHeaderRow; row++) {
		for (let col = bounds.start.col; col <= maxLabelCol; col++) {
			const label = normalizePivotAuditText(cellText(sheet.cells.get(row, col)?.value ?? EMPTY))
			if (label === normalizedFieldName) return col
		}
	}
	return findSavedPivotFieldColumn(sheet, bounds, header.row, header.col, fieldName)
}

function findSavedPivotDataStartRow(
	sheet: Workbook['sheets'][number],
	bounds: RangeRef,
	pivot: PivotTableInfo,
	cache: PivotCacheInfo | undefined,
	header: { row: number; col: number },
): number | undefined {
	if (pivot.rowFields.length === 0) return undefined
	const rowFieldCols = pivot.rowFields
		.map((reference) =>
			findSavedPivotRowFieldColumn(
				sheet,
				bounds,
				header,
				pivot.fields[reference.index]?.name ?? cache?.fields[reference.index]?.name ?? '',
			),
		)
		.filter((col): col is number => col !== undefined)
	if (rowFieldCols.length === 0) return undefined
	for (let row = header.row + 1; row <= bounds.end.row; row++) {
		for (const col of rowFieldCols) {
			const label = normalizePivotAuditText(cellText(sheet.cells.get(row, col)?.value ?? EMPTY))
			if (
				label &&
				label !== 'grand total' &&
				!savedPivotFieldReferenceListMatches(pivot, cache, pivot.rowFields, label)
			) {
				return row
			}
		}
	}
	return undefined
}

function findSavedPivotValueColumn(
	sheet: Workbook['sheets'][number],
	bounds: RangeRef,
	header: { row: number; col: number },
	columnHeaderEndRow: number,
	columnFilters: readonly GetPivotDataFilter[],
	row: number,
	rowFilterCount: number,
): number | undefined {
	if (columnFilters.length === 0) {
		return rowFilterCount === 0 && header.col === bounds.start.col
			? rightmostSavedPivotRowValueCol(sheet, bounds, row)
			: header.col
	}
	const minDataCol = Math.max(bounds.start.col + 1, header.col)
	for (let col = minDataCol; col <= bounds.end.col; col++) {
		if (savedPivotColumnMatchesFilters(sheet, bounds, col, columnHeaderEndRow, columnFilters)) {
			return col
		}
	}
	return undefined
}

function savedPivotColumnMatchesFilters(
	sheet: Workbook['sheets'][number],
	bounds: RangeRef,
	col: number,
	columnHeaderEndRow: number,
	filters: readonly GetPivotDataFilter[],
): boolean {
	if (columnHeaderEndRow < bounds.start.row) return false
	for (const filter of filters) {
		let matched = false
		for (let row = bounds.start.row; row <= columnHeaderEndRow; row++) {
			const value = normalizePivotAuditText(cellText(sheet.cells.get(row, col)?.value ?? EMPTY))
			if (value === normalizePivotAuditText(filter.item)) {
				matched = true
				break
			}
		}
		if (!matched) return false
	}
	return true
}

function savedPivotRowMatchesFilters(
	sheet: Workbook['sheets'][number],
	bounds: RangeRef,
	row: number,
	fieldColumns: readonly {
		readonly filter: GetPivotDataFilter
		readonly col: number | undefined
	}[],
): boolean {
	if (fieldColumns.length === 0) {
		const label = cellText(sheet.cells.get(row, bounds.start.col)?.value ?? EMPTY)
		return normalizeGrandTotalLabel(label) === 'Grand Total'
	}
	for (const entry of fieldColumns) {
		if (entry.col === undefined) return false
		const value = normalizePivotAuditText(cellText(sheet.cells.get(row, entry.col)?.value ?? EMPTY))
		if (value !== normalizePivotAuditText(entry.filter.item)) return false
	}
	return true
}

function rightmostSavedPivotRowValueCol(
	sheet: Workbook['sheets'][number],
	bounds: RangeRef,
	row: number,
): number | undefined {
	for (let col = bounds.end.col; col > bounds.start.col; col--) {
		if ((sheet.cells.get(row, col)?.value ?? EMPTY).kind !== 'empty') return col
	}
	return undefined
}

function buildExternalReferenceUsages(
	workbook: Workbook,
	analysis: WorkbookFormulaAnalysis,
): ExternalReferenceUsageInfo[] {
	const usages: ExternalReferenceUsageInfo[] = []
	for (const formula of analysis.formulas.values()) {
		if (!formula.ast) continue
		pushExternalReferenceUsages(workbook, usages, formula.ast, {
			sourceKind: 'cellFormula',
			sourceRef: `${formula.sheetName}!${indexToColumn(formula.col)}${formula.row + 1}`,
			formula: formula.formula,
		})
	}

	for (const name of workbook.definedNames.list()) {
		const parsed = parseFormula(normalizeFormulaInput(name.formula))
		if (!parsed.ok) continue
		pushExternalReferenceUsages(workbook, usages, parsed.value, {
			sourceKind: 'definedName',
			name: name.name,
			formula: name.formula,
		})
	}

	for (const chart of workbook.chartParts) {
		for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
			const series = chart.series[seriesIndex]
			if (!series) continue
			const sourceRef = `${chart.partPath}#series${seriesIndex}`
			if (series.nameRef) {
				pushParsedExternalReferenceUsages(workbook, usages, series.nameRef, {
					sourceKind: 'chartSeriesName',
					sourceRef,
					formula: series.nameRef,
				})
			}
			if (series.categoryRef) {
				pushParsedExternalReferenceUsages(workbook, usages, series.categoryRef, {
					sourceKind: 'chartSeriesCategory',
					sourceRef,
					formula: series.categoryRef,
				})
			}
			if (series.valueRef) {
				pushParsedExternalReferenceUsages(workbook, usages, series.valueRef, {
					sourceKind: 'chartSeriesValue',
					sourceRef,
					formula: series.valueRef,
				})
			}
		}
	}

	for (const sheet of workbook.sheets) {
		for (const table of sheet.tables) {
			for (const column of table.columns) {
				const sourceRef = `${sheet.name}!${table.name}[${column.name}]`
				if (column.formula) {
					pushParsedExternalReferenceUsages(workbook, usages, column.formula, {
						sourceKind: 'tableColumnFormula',
						sourceRef,
						formula: column.formula,
					})
				}
				if (column.totalsRowFormula) {
					pushParsedExternalReferenceUsages(workbook, usages, column.totalsRowFormula, {
						sourceKind: 'tableTotalsRowFormula',
						sourceRef,
						formula: column.totalsRowFormula,
					})
				}
			}
		}
		for (const format of sheet.conditionalFormats) {
			for (const rule of format.rules) {
				for (const formulaText of conditionalFormatFormulaTexts(rule)) {
					pushParsedExternalReferenceUsages(workbook, usages, formulaText, {
						sourceKind: 'conditionalFormat',
						sourceRef: `${sheet.name}!${format.sqref}`,
						formula: formulaText,
					})
				}
			}
		}
		for (const validation of sheet.dataValidations) {
			for (const formulaText of [validation.formula1, validation.formula2]) {
				if (!formulaText) continue
				pushParsedExternalReferenceUsages(workbook, usages, formulaText, {
					sourceKind: 'dataValidation',
					sourceRef: `${sheet.name}!${validation.sqref}`,
					formula: formulaText,
				})
			}
		}
		for (const format of sheet.x14ConditionalFormats) {
			for (const formulaText of format.formulas) {
				pushParsedExternalReferenceUsages(workbook, usages, formulaText, {
					sourceKind: 'x14ConditionalFormat',
					sourceRef: `${sheet.name}!${format.sqref}`,
					formula: formulaText,
				})
			}
			for (const formulaText of x14ConditionalFormatValueFormulas(format)) {
				pushParsedExternalReferenceUsages(workbook, usages, formulaText, {
					sourceKind: 'x14ConditionalFormat',
					sourceRef: `${sheet.name}!${format.sqref}`,
					formula: formulaText,
				})
			}
		}
		for (const validation of sheet.x14DataValidations) {
			for (const formulaText of [validation.formula1, validation.formula2]) {
				if (!formulaText) continue
				pushParsedExternalReferenceUsages(workbook, usages, formulaText, {
					sourceKind: 'x14DataValidation',
					sourceRef: `${sheet.name}!${validation.sqref}`,
					formula: formulaText,
				})
			}
		}
		for (const group of sheet.sparklineGroups) {
			const groupRef = `${sheet.name}!sparklineGroup${group.groupIndex}`
			if (group.range) {
				pushParsedExternalReferenceUsages(workbook, usages, group.range, {
					sourceKind: 'sparklineGroupRange',
					sourceRef: groupRef,
					formula: group.range,
				})
			}
			if (group.dateAxisRange) {
				pushParsedExternalReferenceUsages(workbook, usages, group.dateAxisRange, {
					sourceKind: 'sparklineDateAxisRange',
					sourceRef: groupRef,
					formula: group.dateAxisRange,
				})
			}
			for (let index = 0; index < (group.sparklines?.length ?? 0); index++) {
				const sparkline = group.sparklines?.[index]
				if (!sparkline?.range) continue
				pushParsedExternalReferenceUsages(workbook, usages, sparkline.range, {
					sourceKind: 'sparklineRange',
					sourceRef: `${groupRef}#sparkline${index}`,
					formula: sparkline.range,
				})
			}
		}
	}
	return usages
}

function pushParsedExternalReferenceUsages(
	workbook: Workbook,
	usages: ExternalReferenceUsageInfo[],
	formulaText: string,
	source: Pick<ExternalReferenceUsageInfo, 'sourceKind' | 'sourceRef' | 'name' | 'formula'>,
): void {
	const parsed = parseFormula(normalizeFormulaInput(formulaText))
	if (!parsed.ok) return
	pushExternalReferenceUsages(workbook, usages, parsed.value, source)
}

function pushExternalReferenceUsages(
	workbook: Workbook,
	usages: ExternalReferenceUsageInfo[],
	ast: FormulaNode,
	source: Pick<ExternalReferenceUsageInfo, 'sourceKind' | 'sourceRef' | 'name' | 'formula'>,
): void {
	for (const group of externalReferenceGroups(collectFormulaReferences(ast))) {
		const externalReference = resolveExternalReference(
			workbook.externalReferenceDetails,
			group.workbook,
		)
		usages.push({
			workbook: group.workbook,
			...(group.sheet ? { sheet: group.sheet } : {}),
			...(group.sheetSpan ? { sheetSpan: group.sheetSpan } : {}),
			...source,
			references: group.references,
			...(externalReference ? { externalReference } : {}),
		})
	}
}

function conditionalFormatFormulaTexts(
	rule: Workbook['sheets'][number]['conditionalFormats'][number]['rules'][number],
): string[] {
	return [
		...rule.formulas,
		...conditionalFormatValueObjectFormulas(rule.colorScale?.cfvo),
		...conditionalFormatValueObjectFormulas(rule.dataBar?.cfvo),
		...conditionalFormatValueObjectFormulas(rule.iconSet?.cfvo),
	]
}

function x14ConditionalFormatValueFormulas(
	format: Workbook['sheets'][number]['x14ConditionalFormats'][number],
): string[] {
	return [
		...conditionalFormatValueObjectFormulas(format.dataBar?.cfvo),
		...conditionalFormatValueObjectFormulas(format.iconSet?.cfvo),
	]
}

function conditionalFormatValueObjectFormulas(
	values: readonly { readonly type?: string; readonly value?: string }[] | undefined,
): string[] {
	return (
		values?.flatMap((value) => (value.type === 'formula' && value.value ? [value.value] : [])) ?? []
	)
}

function resolveExternalReference(
	details: readonly ExternalReferenceInfo[],
	workbookToken: string,
): ExternalReferenceInfo | undefined {
	const numericIndex = parseExternalReferenceIndex(workbookToken)
	if (numericIndex !== undefined) return copyExternalReferenceInfo(details[numericIndex])

	const tokenName = externalReferenceBasename(workbookToken)
	const matches = details.filter((entry) => {
		if (entry.target === workbookToken || entry.partPath === workbookToken) return true
		return entry.target !== undefined && externalReferenceBasename(entry.target) === tokenName
	})
	return matches.length === 1 ? copyExternalReferenceInfo(matches[0]) : undefined
}

function parseExternalReferenceIndex(workbookToken: string): number | undefined {
	if (!/^\d+$/.test(workbookToken)) return undefined
	const index = Number(workbookToken)
	return Number.isSafeInteger(index) && index > 0 ? index - 1 : undefined
}

function externalReferenceBasename(path: string): string {
	const normalized = path.replace(/\\/g, '/')
	return normalized.slice(normalized.lastIndexOf('/') + 1)
}

function copyExternalReferenceInfo(
	reference: ExternalReferenceInfo | undefined,
): ExternalReferenceInfo | undefined {
	return reference ? { ...reference } : undefined
}

function externalReferenceGroups(references: readonly FormulaReferenceInfo[]): Array<{
	workbook: string
	sheet?: string
	sheetSpan?: { readonly startSheet: string; readonly endSheet: string }
	references: string[]
}> {
	const groups = new Map<
		string,
		{
			workbook: string
			sheet?: string
			sheetSpan?: { readonly startSheet: string; readonly endSheet: string }
			references: string[]
		}
	>()
	for (const reference of flattenFormulaReferences(references)) {
		if (reference.scope?.kind !== 'external' && reference.scope?.kind !== 'externalSheetSpan') {
			continue
		}
		const key =
			reference.scope.kind === 'external'
				? `${reference.scope.workbook}\u0000${reference.scope.sheet}`
				: `${reference.scope.workbook}\u0000${reference.scope.startSheet}\u0000${reference.scope.endSheet}`
		const existing = groups.get(key)
		if (existing) {
			existing.references.push(reference.text)
			continue
		}
		groups.set(
			key,
			reference.scope.kind === 'external'
				? {
						workbook: reference.scope.workbook,
						sheet: reference.scope.sheet,
						references: [reference.text],
					}
				: {
						workbook: reference.scope.workbook,
						sheetSpan: {
							startSheet: reference.scope.startSheet,
							endSheet: reference.scope.endSheet,
						},
						references: [reference.text],
					},
		)
	}
	return [...groups.values()].map((group) => ({
		...group,
		references: [...new Set(group.references)],
	}))
}

function flattenFormulaReferences(
	references: readonly FormulaReferenceInfo[],
): FormulaReferenceInfo[] {
	const out: FormulaReferenceInfo[] = []
	for (const reference of references) {
		if (reference.kind === 'union' || reference.kind === 'intersection') {
			out.push(...flattenFormulaReferences(reference.members))
			continue
		}
		out.push(reference)
		if (
			(reference.kind === 'spill' || reference.kind === 'implicitIntersection') &&
			reference.target
		) {
			out.push(...flattenFormulaReferences([reference.target]))
		}
	}
	return out
}

function buildPivotRefreshPlans(
	workbook: Workbook,
	cellsHydrated: boolean,
): PivotRefreshPlanInfo[] {
	return workbook.pivotCaches.map((cache) => {
		const linkedPivots = workbook.pivotTables.filter(
			(pivot) => pivot.cacheId !== undefined && pivot.cacheId === cache.cacheId,
		)
		const pivotTables = linkedPivots.map((pivot) => ({
			partPath: pivot.partPath,
			sheetName: pivot.sheetName,
			...(pivot.name !== undefined ? { name: pivot.name } : {}),
			...(pivot.locationRef !== undefined ? { locationRef: pivot.locationRef } : {}),
		}))
		const cacheState = pivotRefreshCacheState(cache)
		const outputSummary = pivotRefreshOutputSummary(workbook, cellsHydrated, linkedPivots)
		const outputState = pivotRefreshOutputState(cacheState, outputSummary)
		const warnings = pivotRefreshWarnings(cache, pivotTables.length, outputState, outputSummary)
		const recommendedOps = pivotRefreshRecommendedOps(cache, outputState)
		return {
			partPath: cache.partPath,
			...(cache.cacheId !== undefined ? { cacheId: cache.cacheId } : {}),
			...(cache.sourceSheet !== undefined ? { sourceSheet: cache.sourceSheet } : {}),
			...(cache.sourceRef !== undefined ? { sourceRef: cache.sourceRef } : {}),
			...(cache.records
				? {
						cacheRecords: {
							partPath: cache.records.partPath,
							...(cache.records.declaredCount !== undefined
								? { declaredCount: cache.records.declaredCount }
								: {}),
							parsedCount: cache.records.parsedCount,
							...(cache.records.materializedCount !== undefined
								? { materializedCount: cache.records.materializedCount }
								: {}),
							...(cache.records.materializedComplete !== undefined
								? { materializedComplete: cache.records.materializedComplete }
								: {}),
							valueKindCounts: cache.records.valueKindCounts.map((count) => ({ ...count })),
						},
					}
				: {}),
			pivotTables,
			outputState,
			canRefreshHeadlessly: false,
			requiresExternalRefresh: outputState !== 'cached',
			warnings,
			recommendedOps,
		}
	})
}

function pivotRefreshCacheState(cache: PivotCacheInfo): PivotRefreshPlanInfo['outputState'] {
	if (cache.invalid) return 'stale'
	if (cache.refreshOnLoad) return 'refresh-on-open'
	if (cache.saveData === false) return 'not-saved'
	if (cache.recordsPartPath === undefined && cache.recordCount === undefined) return 'unknown'
	return 'cached'
}

type PivotOutputRefreshSummary = {
	readonly notSavedCount: number
	readonly unknownCount: number
}

function pivotRefreshOutputState(
	cacheState: PivotRefreshPlanInfo['outputState'],
	outputSummary: PivotOutputRefreshSummary,
): PivotRefreshPlanInfo['outputState'] {
	if (cacheState !== 'cached') return cacheState
	if (outputSummary.notSavedCount > 0) return 'not-saved'
	if (outputSummary.unknownCount > 0) return 'unknown'
	return 'cached'
}

function pivotRefreshOutputSummary(
	workbook: Workbook,
	cellsHydrated: boolean,
	pivots: readonly PivotTableInfo[],
): PivotOutputRefreshSummary {
	let notSavedCount = 0
	let unknownCount = 0
	for (const pivot of pivots) {
		const state = pivotSavedOutputState(workbook, cellsHydrated, pivot)
		if (state === 'not-saved') notSavedCount++
		if (state === 'unknown') unknownCount++
	}
	return { notSavedCount, unknownCount }
}

function pivotSavedOutputState(
	workbook: Workbook,
	cellsHydrated: boolean,
	pivot: PivotTableInfo,
): 'cached' | 'not-saved' | 'unknown' {
	if (!cellsHydrated) return 'unknown'
	if (isEmptyPivotOutput(pivot)) return 'cached'
	if (!pivot.locationRef) return 'unknown'
	let bounds: RangeRef
	try {
		bounds = parseRange(pivot.locationRef)
	} catch {
		return 'unknown'
	}
	const sheet = workbook.sheets.find((entry) => entry.name === pivot.sheetName)
	if (!sheet) return 'unknown'
	for (let row = bounds.start.row; row <= bounds.end.row; row++) {
		for (let col = bounds.start.col; col <= bounds.end.col; col++) {
			if ((sheet.cells.get(row, col)?.value ?? EMPTY).kind !== 'empty') return 'cached'
		}
	}
	return 'not-saved'
}

function pivotRefreshWarnings(
	cache: PivotCacheInfo,
	pivotTableCount: number,
	outputState: PivotRefreshPlanInfo['outputState'],
	outputSummary: PivotOutputRefreshSummary,
): string[] {
	const warnings: string[] = []
	if (outputState === 'stale') {
		warnings.push('Pivot output cells are stale until the cache is refreshed by Excel.')
	}
	if (outputState === 'refresh-on-open') {
		warnings.push(
			'Workbook requests pivot refresh on open; saved output may differ after Excel opens it.',
		)
	}
	if (outputState === 'not-saved') {
		if (cache.saveData === false) {
			warnings.push(
				'Pivot cache records are not saved; a pivot-aware application must rebuild them.',
			)
		}
	}
	if (outputState === 'unknown') {
		warnings.push('Pivot cache freshness is unknown because cache records were not inventoried.')
	}
	if (outputSummary.notSavedCount > 0) {
		warnings.push(
			`${outputSummary.notSavedCount} PivotTable output range(s) have no saved cells; a pivot-aware application must materialize them.`,
		)
	}
	if (outputSummary.unknownCount > 0) {
		warnings.push(
			`${outputSummary.unknownCount} PivotTable output range(s) could not be inspected for saved cells.`,
		)
	}
	if (pivotTableCount > 0 && (cache.sourceSheet !== undefined || cache.sourceRef !== undefined)) {
		warnings.push(
			'Ascend can edit pivot cache metadata but does not recalculate pivot output cells headlessly.',
		)
	}
	return warnings
}

function pivotRefreshRecommendedOps(
	cache: PivotCacheInfo,
	outputState: PivotRefreshPlanInfo['outputState'],
): PivotRefreshRecommendedOp[] {
	if (outputState === 'cached') return []
	return [
		{
			op: 'setPivotCache',
			...(cache.partPath !== undefined ? { partPath: cache.partPath } : {}),
			...(cache.cacheId !== undefined ? { cacheId: cache.cacheId } : {}),
			refreshOnLoad: true,
			invalid: true,
			saveData: false,
		},
	]
}

function buildWorkbookRefreshMetadata(
	workbook: Workbook,
	report: CompatibilityReport,
	cellsHydrated: boolean,
): WorkbookRefreshMetadataInfo {
	const entries: WorkbookRefreshMetadataEntry[] = []
	const calcWarnings: string[] = []
	const refreshOnOpen =
		workbook.calcSettings.fullCalcOnLoad || workbook.calcSettings.forceFullCalc === true
	if (refreshOnOpen) {
		calcWarnings.push('Workbook requests full recalculation on open.')
	}
	if (workbook.calcSettings.calcMode === 'manual') {
		calcWarnings.push('Workbook is in manual calculation mode.')
	}
	if (calcWarnings.length > 0) {
		entries.push({
			kind: 'calcSettings',
			partPath: 'xl/workbook.xml',
			state: refreshOnOpen ? 'refresh-on-open' : 'manual-calc',
			refreshOnLoad: refreshOnOpen,
			warnings: calcWarnings,
			recommendedOps: [],
		})
	}

	const calcChain = report.features.find((feature) => feature.feature === 'calcChain')
	if (calcChain) {
		for (const partPath of calcChain.locations) {
			entries.push({
				kind: 'calcChain',
				partPath,
				state: 'cached',
				warnings: [
					'Imported calc chain is preserved unless a formula-topology edit requires full recalculation on open.',
				],
				recommendedOps: [],
			})
		}
	}

	for (const plan of buildPivotRefreshPlans(workbook, cellsHydrated)) {
		entries.push({
			kind: 'pivotCache',
			partPath: plan.partPath,
			state: plan.outputState,
			...(plan.cacheId !== undefined ? { cacheId: plan.cacheId } : {}),
			...(plan.sourceSheet !== undefined ? { sourceSheet: plan.sourceSheet } : {}),
			...(plan.sourceRef !== undefined ? { sourceRef: plan.sourceRef } : {}),
			refreshOnLoad: plan.outputState === 'refresh-on-open',
			invalid: plan.outputState === 'stale',
			warnings: plan.warnings,
			recommendedOps: plan.recommendedOps,
		})
	}

	for (const part of workbook.connectionParts) {
		if (part.kind !== 'connection' && part.kind !== 'queryTable') continue
		const state = connectionRefreshState(part)
		entries.push({
			kind: part.kind === 'queryTable' ? 'queryTable' : 'workbookConnection',
			partPath: part.partPath,
			state,
			...(part.name !== undefined ? { name: part.name } : {}),
			...(part.sheetName !== undefined ? { sheetName: part.sheetName } : {}),
			...(part.connectionId !== undefined ? { connectionId: part.connectionId } : {}),
			...(part.refreshOnLoad !== undefined ? { refreshOnLoad: part.refreshOnLoad } : {}),
			...(part.saveData !== undefined ? { saveData: part.saveData } : {}),
			...(part.refreshedVersion !== undefined ? { refreshedVersion: part.refreshedVersion } : {}),
			warnings: connectionRefreshWarnings(part, state),
			recommendedOps: connectionRefreshRecommendedOps(part, state),
		})
	}

	return {
		entries,
		refreshOnOpenCount: entries.filter((entry) => entry.state === 'refresh-on-open').length,
		staleCacheCount: entries.filter((entry) => entry.state === 'stale').length,
		notSavedCount: entries.filter(
			(entry) => entry.state === 'not-saved' || entry.saveData === false,
		).length,
		unknownCount: entries.filter((entry) => entry.state === 'unknown').length,
	}
}

function connectionRefreshState(
	part: WorkbookConnectionPartInfo,
): WorkbookRefreshMetadataEntry['state'] {
	if (part.refreshOnLoad) return 'refresh-on-open'
	if (part.saveData === false) return 'not-saved'
	if (part.refreshedVersion === undefined) return 'unknown'
	return 'cached'
}

function connectionRefreshWarnings(
	part: WorkbookConnectionPartInfo,
	state: WorkbookRefreshMetadataEntry['state'],
): string[] {
	if (state === 'refresh-on-open') {
		return ['Connection requests refresh on open; saved output may change after Excel opens it.']
	}
	if (state === 'not-saved') {
		return ['Connection cache data is not saved; a connection-aware application must refresh it.']
	}
	if (state === 'unknown') {
		return ['Connection freshness is unknown because refresh version metadata is absent.']
	}
	return part.kind === 'queryTable'
		? ['Query table refresh metadata is inspectable and editable without executing the query.']
		: [
				'Workbook connection refresh metadata is inspectable and editable without executing the connection.',
			]
}

function connectionRefreshRecommendedOps(
	part: WorkbookConnectionPartInfo,
	state: WorkbookRefreshMetadataEntry['state'],
): readonly unknown[] {
	if (state === 'cached') return []
	return [
		{
			op: 'setConnectionRefresh',
			partPath: part.partPath,
			...(part.name !== undefined ? { name: part.name } : {}),
			...(part.connectionId !== undefined ? { connectionId: part.connectionId } : {}),
			refreshOnLoad: true,
			saveData: false,
		},
	]
}

function budgetAgentViewResult(
	result: AgentViewResult,
	requestedApproxTokens: number,
): AgentViewResult {
	const budget = Math.max(64, Math.floor(requestedApproxTokens))
	const unbudgetedApproxTokens = estimateAgentViewApproxTokens(result)
	const baseColumnSampleValues = countAgentViewColumnSampleValues(result.columns)
	let sampleRowLimit = result.samples.length
	let columnSampleValueLimit = result.columns.reduce(
		(max, column) => Math.max(max, column.sampleValues.length),
		0,
	)
	let formulaPatternLimit = result.formulaPatterns.length
	let candidate = buildBudgetedAgentViewResult(
		result,
		budget,
		sampleRowLimit,
		columnSampleValueLimit,
		formulaPatternLimit,
		baseColumnSampleValues,
		unbudgetedApproxTokens,
	)

	while (
		candidate.budget &&
		candidate.budget.estimatedApproxTokens > budget &&
		(sampleRowLimit > 0 || columnSampleValueLimit > 0 || formulaPatternLimit > 0)
	) {
		if (sampleRowLimit > 2) sampleRowLimit--
		else if (columnSampleValueLimit > 2) columnSampleValueLimit--
		else if (formulaPatternLimit > 6) formulaPatternLimit--
		else if (sampleRowLimit > 0) sampleRowLimit--
		else if (columnSampleValueLimit > 0) columnSampleValueLimit--
		else formulaPatternLimit--
		candidate = buildBudgetedAgentViewResult(
			result,
			budget,
			sampleRowLimit,
			columnSampleValueLimit,
			formulaPatternLimit,
			baseColumnSampleValues,
			unbudgetedApproxTokens,
		)
	}

	return candidate
}

function buildBudgetedAgentViewResult(
	result: AgentViewResult,
	requestedApproxTokens: number,
	sampleRowLimit: number,
	columnSampleValueLimit: number,
	formulaPatternLimit: number,
	baseColumnSampleValues: number,
	unbudgetedApproxTokens: number,
): AgentViewResult {
	const columns = result.columns.map((column) =>
		column.sampleValues.length > columnSampleValueLimit
			? { ...column, sampleValues: column.sampleValues.slice(0, columnSampleValueLimit) }
			: column,
	)
	const budgeted = {
		...result,
		formulaPatterns: result.formulaPatterns.slice(0, formulaPatternLimit),
		columns,
		samples: result.samples.slice(0, sampleRowLimit),
	}
	const omittedSampleRows = result.samples.length - budgeted.samples.length
	const omittedColumnSampleValues =
		baseColumnSampleValues - countAgentViewColumnSampleValues(columns)
	const omittedFormulaPatterns = result.formulaPatterns.length - budgeted.formulaPatterns.length
	const omittedEvidence = buildAgentViewOmittedEvidence(
		result,
		budgeted,
		formulaPatternLimit,
		omittedSampleRows,
		omittedColumnSampleValues,
		omittedFormulaPatterns,
	)
	const budget = {
		requestedApproxTokens,
		estimatedApproxTokens: 0,
		unbudgetedApproxTokens,
		estimator: 'json-bytes-div-4' as const,
		truncated: omittedSampleRows > 0 || omittedColumnSampleValues > 0 || omittedFormulaPatterns > 0,
		omittedSampleRows,
		omittedColumnSampleValues,
		omittedFormulaPatterns,
		...(omittedEvidence ? { omittedEvidence } : {}),
	}
	const withBudget = { ...budgeted, budget }
	return {
		...withBudget,
		budget: {
			...budget,
			estimatedApproxTokens: estimateAgentViewApproxTokens(withBudget),
		},
	}
}

const MAX_OMITTED_COLUMN_SAMPLE_HINTS = 2

function buildAgentViewOmittedEvidence(
	result: AgentViewResult,
	budgeted: Omit<AgentViewResult, 'budget'>,
	formulaPatternLimit: number,
	omittedSampleRows: number,
	omittedColumnSampleValues: number,
	omittedFormulaPatterns: number,
): NonNullable<import('./types.ts').AgentViewBudgetInfo['omittedEvidence']> | undefined {
	if (omittedSampleRows + omittedColumnSampleValues + omittedFormulaPatterns === 0) return undefined
	const omittedRows = result.samples.slice(budgeted.samples.length).map((sample) => sample.row)
	const columnHints = result.columns
		.map((column, index) => {
			const shown = budgeted.columns[index]?.sampleValues.length ?? 0
			const omittedValues = Math.max(0, column.sampleValues.length - shown)
			return omittedValues > 0 ? { col: column.col, ref: column.ref, omittedValues } : undefined
		})
		.filter(
			(
				entry,
			): entry is { readonly col: number; readonly ref: string; readonly omittedValues: number } =>
				entry !== undefined,
		)
	const omittedPatterns = result.formulaPatterns.slice(formulaPatternLimit)
	return {
		...(omittedRows.length > 0
			? {
					sampleRows: {
						count: omittedRows.length,
						firstRow: omittedRows[0] ?? 0,
						lastRow: omittedRows[omittedRows.length - 1] ?? 0,
						remainingRows: 0,
					},
				}
			: {}),
		...(columnHints.length > 0
			? {
					columnSampleValues: {
						omittedValues: omittedColumnSampleValues,
						columns: columnHints.slice(0, MAX_OMITTED_COLUMN_SAMPLE_HINTS),
						remainingColumns: Math.max(0, columnHints.length - MAX_OMITTED_COLUMN_SAMPLE_HINTS),
					},
				}
			: {}),
		...(omittedPatterns.length > 0
			? {
					formulaPatterns: {
						count: omittedFormulaPatterns,
						firstOmittedIndex: formulaPatternLimit,
						...(omittedPatterns[0]?.examples[0]
							? { nextExampleRef: omittedPatterns[0].examples[0] }
							: {}),
					},
				}
			: {}),
	}
}

function countAgentViewColumnSampleValues(columns: readonly AgentColumnSummary[]): number {
	return columns.reduce((total, column) => total + column.sampleValues.length, 0)
}

function estimateAgentViewApproxTokens(value: unknown): number {
	return Math.ceil(JSON.stringify(value).length / 4)
}

function flattenForAgent(value: import('@ascend/schema').CellValue): FlatCellValue {
	switch (value.kind) {
		case 'empty':
			return null
		case 'number':
			return value.value
		case 'string':
			return value.value
		case 'boolean':
			return value.value
		case 'date':
			return value.serial
		case 'error':
			return value.value
		case 'richText':
			return value.runs.map((run) => run.text).join('')
		case 'array':
			return null
	}
}

function groupCompactCellsByRow(cells: readonly CompactCellInfo[]): Map<number, CompactCellInfo[]> {
	const rows = new Map<number, CompactCellInfo[]>()
	for (let i = 0; i < cells.length; i++) {
		const cell = cells[i] as CompactCellInfo
		const bucket = rows.get(cell.row)
		if (bucket) bucket.push(cell)
		else rows.set(cell.row, [cell])
	}
	for (const bucket of rows.values()) {
		bucket.sort((left, right) => left.col - right.col)
	}
	return rows
}

function extractFormulaFunctions(formula: string): string[] {
	const matches = formula.toUpperCase().match(/[A-Z][A-Z0-9._]*\s*\(/g) ?? []
	const names = new Set<string>()
	for (let i = 0; i < matches.length; i++) {
		const match = matches[i]
		if (!match) continue
		names.add(match.replace(/\s*\($/, ''))
	}
	return [...names]
}

function normalizeFormulaPattern(formula: string): string {
	return formula
		.toUpperCase()
		.replace(/\$?[A-Z]{1,3}\$?\d+/g, 'REF')
		.replace(/\d+(\.\d+)?/g, '#')
}
