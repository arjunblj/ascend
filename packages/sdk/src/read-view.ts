import { parseA1, type RangeRef, type Workbook } from '@ascend/core'
import {
	analyzeWorkbook,
	analyzeWorkbookDependencies,
	analyzeWorkbookFormulas,
	type CellKey,
	cellKey,
	createSnapshot,
	diffWorkbooks,
	resolveCellFormulaText,
	type WorkbookAnalysis,
	type WorkbookDependencyAnalysis,
	type WorkbookDiff,
	type WorkbookFormulaAnalysis,
	type WorkbookSnapshot,
} from '@ascend/engine'
import { parseFormula, printFormula } from '@ascend/formulas'
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
import { SheetHandle } from './sheet-handle.ts'
import { TableHandle } from './table-handle.ts'
import type {
	CompactRangeInfo,
	CompactRangeWindowInfo,
	DefinedNameInfo,
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
			pivotTableCount: this.wb.pivotTables.length,
			pivotCacheCount: this.wb.pivotCaches.length,
			slicerCount: this.wb.slicers.length,
			slicerCacheCount: this.wb.slicerCaches.length,
			sourceFormat: this.compat.sourceFormat,
			workbookViewCount: this.wb.workbookViews.length,
			externalReferenceCount: this.wb.externalReferences.length,
			workbookViews: this.wb.workbookViews.map((view) => ({ ...view })),
			externalReferences: [...this.wb.externalReferences],
			hasWorkbookProtection: this.wb.workbookProtection !== null,
			pivotTables: this.wb.pivotTables.map((entry) => ({ ...entry })),
			pivotCaches: this.wb.pivotCaches.map((entry) => ({ ...entry })),
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
		opts?: { includeRefs?: boolean },
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
		opts?: { rowOffset?: number; rowLimit?: number; includeRefs?: boolean },
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
			.map((entry) => ({ ...entry }))
	}

	pivotCaches(): readonly PivotCacheInfo[] {
		return this.wb.pivotCaches.map((entry) => ({ ...entry }))
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

	trace(cellRef: string, opts?: { maxDepth?: number }): TraceResult | undefined {
		if (this.dependencyVerificationIssue()) return undefined
		const { sheetName, ref } = parseFullRef(cellRef, this.wb)
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

	formula(cellRef: string): FormulaInfo | undefined {
		if (this.formulaInfoCache.has(cellRef)) return this.formulaInfoCache.get(cellRef)
		const { sheetName, ref } = parseFullRef(cellRef, this.wb)
		const formulaKey = makeFormulaKey(this.wb, sheetName, ref)
		const analyzed =
			formulaKey !== undefined ? this.formulaAnalysis().formulas.get(formulaKey) : undefined
		const cell = this.sheet(sheetName)?.cell(ref)
		if (!cell || !analyzed) {
			this.formulaInfoCache.set(cellRef, undefined)
			return undefined
		}

		const formula = normalizeFormulaInput(analyzed.formula)
		const tokens = tokenizeFormulaInput(formula)
		if (!analyzed.ast) {
			const info = buildFormulaInfo({
				ref: `${sheetName}!${ref}`,
				formula,
				value: cell.value,
				...(cell.formulaBinding ? { binding: cell.formulaBinding } : {}),
				tokens,
				normalizedFormula: formula,
				functions: [],
				volatile: analyzed.volatile,
				...(analyzed.parseError ? { parseError: analyzed.parseError } : {}),
			})
			this.formulaInfoCache.set(cellRef, info)
			return info
		}

		const info = buildFormulaInfo({
			ref: `${sheetName}!${ref}`,
			formula,
			value: cell.value,
			...(cell.formulaBinding ? { binding: cell.formulaBinding } : {}),
			tokens,
			ast: analyzed.ast,
			normalizedFormula: printFormula(analyzed.ast),
			functions: [...collectFunctionNames(analyzed.ast)],
			volatile: analyzed.volatile,
		})
		this.formulaInfoCache.set(cellRef, info)
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

function parseFullRef(cellRef: string, workbook: Workbook): { sheetName: string; ref: string } {
	const bang = cellRef.indexOf('!')
	if (bang !== -1) {
		const sheetName = cellRef.substring(0, bang).replace(/^'|'$/g, '')
		return { sheetName, ref: cellRef.substring(bang + 1) }
	}
	const firstSheet = workbook.sheets[0]
	const sheetName = firstSheet ? firstSheet.name : 'Sheet1'
	return { sheetName, ref: cellRef }
}

function normalizeFormulaInput(formula: string): string {
	return formula.startsWith('=') ? formula.slice(1) : formula
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

function makeFormulaKey(workbook: Workbook, sheetName: string, ref: string): CellKey | undefined {
	const sheetIndex = workbook.sheets.findIndex((sheet) => sheet.name === sheetName)
	if (sheetIndex === -1) return undefined
	const cellRef = parseA1(ref)
	return cellKey(sheetIndex, cellRef.row, cellRef.col)
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
