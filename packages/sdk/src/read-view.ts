import {
	type ChartPartInfo,
	indexToColumn,
	parseA1,
	type RangeRef,
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
import { normalizeFormulaInput, parseFormula, printFormula } from '@ascend/formulas'
import type { CompatibilityReport } from '@ascend/schema'
import { EMPTY } from '@ascend/schema'
import { trace as verifyTrace } from '@ascend/verify'
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
	AgentColumnSummary,
	AgentReadOptions,
	AgentSampleRow,
	AgentViewOptions,
	AgentViewResult,
	CompactCellInfo,
	CompactRangeInfo,
	CompactRangeWindowInfo,
	DefinedNameInfo,
	FlatCellValue,
	FormulaInfo,
	PivotCacheInfo,
	PivotTableInfo,
	RangeObjectsInfo,
	RangeRowsInfo,
	RangeWindowInfo,
	SheetInfo,
	SheetInspectInfo,
	SlicerCacheInfo,
	SlicerInfo,
	TableInfo,
	TraceResult,
	WorkbookInfo,
	WorkbookLoadInfo,
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
		let totalConditionalFormats = 0
		let totalDataValidations = 0
		let totalImages = 0
		const sheets = this.wb.sheets.map((sheet) => {
			const cellsHydrated = this.loadInfo.cellsHydrated
			const richSheetMetadataHydrated = this.loadInfo.richSheetMetadataHydrated
			const used = cellsHydrated ? sheet.cells.usedRange() : null
			const count = cellsHydrated ? sheet.cells.cellCount() : null
			if (count !== null) totalCells += count
			if (richSheetMetadataHydrated) {
				totalComments += sheet.comments.size
				totalConditionalFormats += sheet.conditionalFormats.length
				totalDataValidations += sheet.dataValidations.length
				totalImages += sheet.imageRefs.length
			}
			return buildSheetInfo(sheet, cellsHydrated, richSheetMetadataHydrated, used, count)
		})
		const info = {
			sheetCount: this.loadInfo.sourceSheets.length,
			loadedSheetCount: this.loadInfo.loadedSheets.length,
			sheets,
			definedNames: this.wb.definedNames.workbookKeys(),
			definedNameDetails: this.definedNames(),
			cellCount: this.loadInfo.cellsHydrated ? totalCells : null,
			commentCount: this.loadInfo.richSheetMetadataHydrated ? totalComments : null,
			conditionalFormatCount: this.loadInfo.richSheetMetadataHydrated
				? totalConditionalFormats
				: null,
			dataValidationCount: this.loadInfo.richSheetMetadataHydrated ? totalDataValidations : null,
			imageCount: this.loadInfo.richSheetMetadataHydrated ? totalImages : null,
			chartCount: this.wb.chartParts.length,
			pivotTableCount: this.wb.pivotTables.length,
			pivotCacheCount: this.wb.pivotCaches.length,
			slicerCount: this.wb.slicers.length,
			slicerCacheCount: this.wb.slicerCaches.length,
			sourceFormat: this.compat.sourceFormat,
			workbookViewCount: this.wb.workbookViews.length,
			externalReferenceCount: this.wb.externalReferences.length,
			workbookViews: this.wb.workbookViews.map((view) => ({ ...view })),
			externalReferences: [...this.wb.externalReferences],
			externalReferenceDetails: this.wb.externalReferenceDetails.map((entry) => ({
				...entry,
			})),
			charts: this.wb.chartParts.map(copyChartInfo),
			hasWorkbookProtection: this.wb.workbookProtection !== null,
			pivotTables: this.wb.pivotTables.map(copyPivotTableInfo),
			pivotCaches: this.wb.pivotCaches.map(copyPivotCacheInfo),
			slicerCaches: this.wb.slicerCaches.map((entry) => ({
				...entry,
				pivotTableNames: [...entry.pivotTableNames],
			})),
			slicers: this.wb.slicers.map((entry) => ({ ...entry })),
			styleSummary: { ...this.wb.styleMetadata },
			themeSummary: {
				hasThemePart: this.wb.preservedTheme !== null,
				...this.wb.themeMetadata,
			},
			compatibility: this.compat,
			load: this.loadInfo,
		}
		this.workbookInfoCache = info
		return info
	}

	visualInventory(): WorkbookVisualInventoryInfo {
		let totalImages = 0
		const charts = this.wb.chartParts.map(copyChartInfo)
		const sheets = this.wb.sheets.map((sheet) => {
			const drawingRefs = this.loadInfo.cellsHydrated ? { ...sheet.drawingRefs } : null
			const imageRefs = this.loadInfo.richSheetMetadataHydrated ? [...sheet.imageRefs] : null
			if (imageRefs) totalImages += imageRefs.length
			return {
				sheet: sheet.name,
				drawingRefs,
				hasDrawing: drawingRefs ? drawingRefs.hasDrawing : null,
				hasLegacyDrawing: drawingRefs ? drawingRefs.hasLegacyDrawing : null,
				imageRefs,
				imageCount: imageRefs ? imageRefs.length : null,
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
			notes.push('Image references require rich sheet metadata hydration.')
		if (charts.length > 0)
			notes.push('Chart parts expose type, title, and series source refs; source edits are staged.')
		else if (packageChartFeatureCount > 0)
			notes.push('Chart parts are preserved but not structurally parsed in this load mode.')
		if (packageDrawingFeatureCount > 0)
			notes.push(
				'Drawing and shape parts are currently preserve-first except parsed image anchors.',
			)
		return {
			load: this.loadInfo,
			packageFeatures,
			sheets,
			sheetImageCount: this.loadInfo.richSheetMetadataHydrated ? totalImages : null,
			charts,
			structuredChartCount: charts.length,
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
			hyperlinks: richSheetMetadataHydrated
				? [...sheet.hyperlinks.entries()].map(([ref, hyperlink]) => ({ ref, ...hyperlink }))
				: null,
			ignoredErrors: cellsHydrated ? [...sheet.ignoredErrors] : null,
			conditionalFormats: richSheetMetadataHydrated ? [...sheet.conditionalFormats] : null,
			dataValidations: richSheetMetadataHydrated ? [...sheet.dataValidations] : null,
			imageRefs: richSheetMetadataHydrated ? [...sheet.imageRefs] : null,
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

	slicerCaches(): readonly SlicerCacheInfo[] {
		return this.wb.slicerCaches.map((entry) => ({
			...entry,
			pivotTableNames: [...entry.pivotTableNames],
		}))
	}

	slicers(): readonly SlicerInfo[] {
		return this.wb.slicers.map((entry) => ({ ...entry }))
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
		conditionalFormatCount: richSheetMetadataHydrated ? sheet.conditionalFormats.length : null,
		dataValidationCount: richSheetMetadataHydrated ? sheet.dataValidations.length : null,
		hasFrozenPanes: cellsHydrated ? sheet.frozenRows > 0 || sheet.frozenCols > 0 : null,
		colWidthCount: cellsHydrated ? sheet.colWidths.size : null,
		imageCount: richSheetMetadataHydrated ? sheet.imageRefs.length : null,
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

function copyPivotCacheInfo(cache: PivotCacheInfo): PivotCacheInfo {
	return {
		...cache,
		fields: cache.fields.map((field) => ({ ...field })),
	}
}

function copyPivotTableInfo(pivot: PivotTableInfo): PivotTableInfo {
	return {
		...pivot,
		fields: pivot.fields.map((field) => ({ ...field })),
		rowFields: pivot.rowFields.map((field) => ({ ...field })),
		columnFields: pivot.columnFields.map((field) => ({ ...field })),
		pageFields: pivot.pageFields.map((field) => ({ ...field })),
		dataFields: pivot.dataFields.map((field) => ({ ...field })),
	}
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
