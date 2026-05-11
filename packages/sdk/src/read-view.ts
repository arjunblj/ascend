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
	type RangeRef,
	type SheetDrawingObjectRef,
	type SheetImageAnchor,
	type SheetImageRef,
	type Workbook,
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
import type { CellValue, CompatibilityReport } from '@ascend/schema'
import { EMPTY } from '@ascend/schema'
import { trace as verifyTrace } from '@ascend/verify'
import { getCapability, isCapabilityGap } from './capabilities.ts'
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
	CompactCellInfo,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	DefinedNameInfo,
	ExternalReferenceInfo,
	ExternalReferenceUsageInfo,
	FlatCellValue,
	FormulaInfo,
	FormulaReferenceInfo,
	GetPivotDataQuery,
	GetPivotDataResult,
	PivotCacheDecodedValueInfo,
	PivotCacheInfo,
	PivotCacheMaterializedRowInfo,
	PivotCacheRecordValueInfo,
	PivotCacheRowsOptions,
	PivotOutputAuditInfo,
	PivotOutputAuditMismatchInfo,
	PivotRefreshPlanInfo,
	PivotRefreshRecommendedOp,
	PivotTableInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
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
	WorkbookRefreshMetadataEntry,
	WorkbookRefreshMetadataInfo,
	WorkbookVisualInventoryInfo,
} from './types.ts'

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
			pivotRefreshPlans: buildPivotRefreshPlans(this.wb.pivotCaches, this.wb.pivotTables),
			refreshMetadata: buildWorkbookRefreshMetadata(this.wb, this.compat),
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
			packageDrawingFeatureCount,
			packageMediaFeatureCount,
			hasPreservedCharts: packageFeatures.some(
				(feature) => feature.category === 'chart' && feature.tier === 'preserved',
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

		return {
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
				.map(([pattern, count]) => ({ pattern, count })),
			columns,
			samples,
			notes,
		}
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

	pivotRefreshPlans(): readonly PivotRefreshPlanInfo[] {
		return buildPivotRefreshPlans(this.wb.pivotCaches, this.wb.pivotTables)
	}

	refreshMetadata(): WorkbookRefreshMetadataInfo {
		return buildWorkbookRefreshMetadata(this.wb, this.compat)
	}

	getPivotData(query: GetPivotDataQuery): GetPivotDataResult {
		return buildGetPivotDataResult(query, this.wb.pivotTables)
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
		rowCount: table.ref.end.row - table.ref.start.row + 1 - headerOffset - totalOffset,
		hasHeaders: table.hasHeaders,
		hasTotals: table.hasTotals,
		autoFilter: table.autoFilter ?? null,
		...(table.sortState ? { sortState: table.sortState } : {}),
		...(table.tableStyleInfo ? { styleInfo: table.tableStyleInfo } : {}),
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
	return workbook.pivotTables.map((pivot) => {
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
		if (!pivot.locationRef)
			return unsupportedPivotAudit(base, 'Pivot table has no output location.')
		const cache = workbook.pivotCaches.find((entry) => entry.cacheId === pivot.cacheId)
		if (!cache) return unsupportedPivotAudit(base, 'Pivot cache metadata was not found.')
		if (!cache.records?.materializedComplete || !cache.records.materializedRecords) {
			return unsupportedPivotAudit(base, 'Pivot cache records are not fully materialized.')
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
		const expected = aggregateSimplePivotOutput(cache, pivot, rowField.index)
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
	})
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

function aggregateSimplePivotOutput(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
	rowFieldIndex: number,
):
	| {
			ok: true
			value: {
				dataFieldNames: readonly string[]
				checkedValueCount: number
				values: ReadonlyMap<string, ReadonlyMap<string, number>>
			}
	  }
	| { ok: false; warning: string } {
	const pageFilters = buildSimplePivotPageFilters(cache, pivot)
	if (!pageFilters.ok) return pageFilters
	const dataFieldNames = pivot.dataFields.map(
		(field, index) => field.name ?? `DataField${index + 1}`,
	)
	const output = new Map<string, Map<string, number>>()
	const baseTotals = new Map<string, Map<string, number>>()
	const includeGrandTotal = pivot.options?.rowGrandTotals !== false
	for (const row of buildPivotCacheRows([cache], {})) {
		if (!pivotCacheRowMatchesPageFilters(row, pageFilters.value)) continue
		const rowLabel = row.values.find((value) => value.fieldIndex === rowFieldIndex)?.value
		if (rowLabel === undefined) continue
		addPivotBaseTotals(cache, row, baseTotals, rowLabel)
		if (includeGrandTotal) addPivotBaseTotals(cache, row, baseTotals, 'Grand Total')
		for (let i = 0; i < pivot.dataFields.length; i++) {
			const dataField = pivot.dataFields[i]
			const dataFieldName = dataFieldNames[i] ?? `DataField${i + 1}`
			if (!dataField) continue
			const field = cache.fields[dataField.fieldIndex]
			if (field?.formula) continue
			const measured = measurePivotDataField(cache, row, dataField)
			if (!measured.ok) return measured
			addPivotOutput(output, rowLabel, dataFieldName, measured.value)
			if (includeGrandTotal) addPivotOutput(output, 'Grand Total', dataFieldName, measured.value)
		}
	}
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
			dataFieldNames,
			checkedValueCount: output.size * dataFieldNames.length,
			values: output,
		},
	}
}

function buildSimplePivotPageFilters(
	cache: PivotCacheInfo,
	pivot: PivotTableInfo,
): { ok: true; value: ReadonlyMap<number, string> } | { ok: false; warning: string } {
	const filters = new Map<number, string>()
	for (const pageField of pivot.pageFields) {
		if (pageField.index < 0) {
			return { ok: false, warning: 'Data-field page filters are not audited.' }
		}
		if (pageField.item === undefined) {
			return { ok: false, warning: 'Multi-select or unset page filters are not audited.' }
		}
		const field = pivot.fields[pageField.index]
		const item = field?.items?.[pageField.item]
		if (!item || item.hidden || item.missing || item.cacheIndex === undefined) {
			return { ok: false, warning: 'Pivot page filter selected item was not resolved.' }
		}
		const sharedItem = cache.fields[pageField.index]?.sharedItems?.find(
			(entry) => entry.index === item.cacheIndex,
		)
		const value = sharedItem?.value
		if (value === undefined) {
			return { ok: false, warning: 'Pivot page filter selected cache item was not resolved.' }
		}
		filters.set(pageField.index, value)
	}
	return { ok: true, value: filters }
}

function pivotCacheRowMatchesPageFilters(
	row: PivotCacheMaterializedRowInfo,
	filters: ReadonlyMap<number, string>,
): boolean {
	for (const [fieldIndex, expected] of filters) {
		const actual = row.values.find((value) => value.fieldIndex === fieldIndex)?.value
		if (actual !== expected) return false
	}
	return true
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
		return { ok: true, value: row.values.some((value) => value.fieldIndex === field.index) ? 1 : 0 }
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
	const headers = new Map<string, { row: number; col: number }>()
	for (let row = bounds.start.row; row <= bounds.end.row; row++) {
		for (let col = bounds.start.col; col <= bounds.end.col; col++) {
			const text = cellText(sheet.cells.get(row, col)?.value ?? EMPTY)
			for (const fieldName of dataFieldNames) {
				if (normalizePivotAuditText(text) === normalizePivotAuditText(fieldName)) {
					headers.set(fieldName, { row, col })
				}
			}
		}
	}
	if (headers.size !== dataFieldNames.length) {
		return { ok: false, warning: 'Pivot output data-field headers were not found.' }
	}
	const headerRow = Math.min(...Array.from(headers.values(), (entry) => entry.row))
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
			const value = sheet.cells.get(row, header.col)?.value ?? EMPTY
			byField.set(fieldName, {
				ref: `${indexToColumn(header.col)}${row + 1}`,
				value,
			})
		}
		output.set(label, byField)
	}
	return { ok: true, value: output }
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
	return normalized === 'grand total' || normalized === 'общий итог' ? 'Grand Total' : label
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

function buildGetPivotDataResult(
	query: GetPivotDataQuery,
	pivots: readonly PivotTableInfo[],
): GetPivotDataResult {
	const normalizedDataField = query.dataField.toLowerCase()
	const normalizedPivotTable = query.pivotTable?.toLowerCase()
	const filters = query.filters ?? []
	const matches = pivots.flatMap((pivot) => {
		if (normalizedPivotTable && pivot.name?.toLowerCase() !== normalizedPivotTable) return []
		const dataField = pivot.dataFields.find((field) => {
			return (
				field.name?.toLowerCase() === normalizedDataField ||
				field.subtotal?.toLowerCase() === normalizedDataField
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
		return [
			{
				pivotTable: copyPivotTableInfo(pivot),
				dataField: { ...dataField },
				matchedFilters,
				unmatchedFilters,
			},
		]
	})
	const warnings = [
		'GETPIVOTDATA metadata can be resolved, but pivot output values are not recalculated headlessly.',
	]
	if (matches.length === 0) {
		warnings.push('No matching pivot table/data field metadata was found.')
	}
	if (matches.some((match) => match.unmatchedFilters.length > 0)) {
		warnings.push(
			'Some requested field/item filters are not present in inspectable pivot metadata.',
		)
	}
	return {
		query,
		matches,
		canResolveOutput: false,
		warnings,
	}
}

function buildExternalReferenceUsages(
	workbook: Workbook,
	analysis: WorkbookFormulaAnalysis,
): ExternalReferenceUsageInfo[] {
	const usages: ExternalReferenceUsageInfo[] = []
	for (const formula of analysis.formulas.values()) {
		if (!formula.ast) continue
		for (const group of externalReferenceGroups(collectFormulaReferences(formula.ast))) {
			const externalReference = resolveExternalReference(
				workbook.externalReferenceDetails,
				group.workbook,
			)
			usages.push({
				workbook: group.workbook,
				...(group.sheet ? { sheet: group.sheet } : {}),
				sourceKind: 'cellFormula',
				sourceRef: `${formula.sheetName}!${indexToColumn(formula.col)}${formula.row + 1}`,
				formula: formula.formula,
				references: group.references,
				...(externalReference ? { externalReference } : {}),
			})
		}
	}

	for (const name of workbook.definedNames.list()) {
		const parsed = parseFormula(normalizeFormulaInput(name.formula))
		if (!parsed.ok) continue
		for (const group of externalReferenceGroups(collectFormulaReferences(parsed.value))) {
			const externalReference = resolveExternalReference(
				workbook.externalReferenceDetails,
				group.workbook,
			)
			usages.push({
				workbook: group.workbook,
				...(group.sheet ? { sheet: group.sheet } : {}),
				sourceKind: 'definedName',
				name: name.name,
				formula: name.formula,
				references: group.references,
				...(externalReference ? { externalReference } : {}),
			})
		}
	}
	return usages
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

function externalReferenceGroups(
	references: readonly FormulaReferenceInfo[],
): Array<{ workbook: string; sheet?: string; references: string[] }> {
	const groups = new Map<string, { workbook: string; sheet?: string; references: string[] }>()
	for (const reference of flattenFormulaReferences(references)) {
		if (reference.scope?.kind !== 'external') continue
		const key = `${reference.scope.workbook}\u0000${reference.scope.sheet}`
		const existing = groups.get(key)
		if (existing) {
			existing.references.push(reference.text)
			continue
		}
		groups.set(key, {
			workbook: reference.scope.workbook,
			sheet: reference.scope.sheet,
			references: [reference.text],
		})
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
	caches: readonly PivotCacheInfo[],
	pivots: readonly PivotTableInfo[],
): PivotRefreshPlanInfo[] {
	return caches.map((cache) => {
		const pivotTables = pivots
			.filter((pivot) => pivot.cacheId !== undefined && pivot.cacheId === cache.cacheId)
			.map((pivot) => ({
				partPath: pivot.partPath,
				sheetName: pivot.sheetName,
				...(pivot.name !== undefined ? { name: pivot.name } : {}),
				...(pivot.locationRef !== undefined ? { locationRef: pivot.locationRef } : {}),
			}))
		const outputState = pivotRefreshOutputState(cache)
		const warnings = pivotRefreshWarnings(cache, pivotTables.length, outputState)
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

function pivotRefreshOutputState(cache: PivotCacheInfo): PivotRefreshPlanInfo['outputState'] {
	if (cache.invalid) return 'stale'
	if (cache.refreshOnLoad) return 'refresh-on-open'
	if (cache.saveData === false) return 'not-saved'
	if (cache.recordsPartPath === undefined && cache.recordCount === undefined) return 'unknown'
	return 'cached'
}

function pivotRefreshWarnings(
	cache: PivotCacheInfo,
	pivotTableCount: number,
	outputState: PivotRefreshPlanInfo['outputState'],
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
		warnings.push('Pivot cache records are not saved; a pivot-aware application must rebuild them.')
	}
	if (outputState === 'unknown') {
		warnings.push('Pivot cache freshness is unknown because cache records were not inventoried.')
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

	for (const plan of buildPivotRefreshPlans(workbook.pivotCaches, workbook.pivotTables)) {
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
