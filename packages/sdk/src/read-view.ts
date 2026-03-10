import { indexToColumn, type RangeRef, type Workbook } from '@ascend/core'
import {
	createSnapshot,
	diffWorkbooks,
	type WorkbookDiff,
	type WorkbookSnapshot,
} from '@ascend/engine'
import {
	extractRefs,
	type FormulaCellRef,
	type FormulaNode,
	functionRegistry,
	parseFormula,
	printFormula,
	tokenize,
} from '@ascend/formulas'
import type { CompatibilityReport } from '@ascend/schema'
import { EMPTY } from '@ascend/schema'
import { trace as verifyTrace } from '@ascend/verify'
import { SheetHandle } from './sheet-handle.ts'
import { TableHandle } from './table-handle.ts'
import type {
	DefinedNameInfo,
	FormulaInfo,
	PivotCacheInfo,
	PivotTableInfo,
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

export class WorkbookReadView {
	protected wb: Workbook
	protected readonly compat: CompatibilityReport
	protected readonly loadInfo: WorkbookLoadInfo
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
			const isHydrated = this.loadInfo.cellsHydrated
			const used = isHydrated ? sheet.cells.usedRange() : null
			const count = isHydrated ? sheet.cells.cellCount() : null
			if (count !== null) totalCells += count
			if (isHydrated) {
				totalComments += sheet.comments.size
				totalConditionalFormats += sheet.conditionalFormats.length
				totalDataValidations += sheet.dataValidations.length
				totalImages += sheet.imageRefs.length
			}
			return buildSheetInfo(sheet, isHydrated, used, count)
		})
		const info = {
			sheetCount: this.loadInfo.sourceSheets.length,
			loadedSheetCount: this.loadInfo.loadedSheets.length,
			sheets,
			definedNames: this.wb.definedNames.workbookKeys(),
			cellCount: this.loadInfo.cellsHydrated ? totalCells : null,
			commentCount: this.loadInfo.cellsHydrated ? totalComments : null,
			conditionalFormatCount: this.loadInfo.cellsHydrated ? totalConditionalFormats : null,
			dataValidationCount: this.loadInfo.cellsHydrated ? totalDataValidations : null,
			imageCount: this.loadInfo.cellsHydrated ? totalImages : null,
			pivotTableCount: this.wb.pivotTables.length,
			pivotCacheCount: this.wb.pivotCaches.length,
			slicerCount: this.wb.slicers.length,
			slicerCacheCount: this.wb.slicerCaches.length,
			sourceFormat: this.compat.sourceFormat,
			workbookViewCount: this.wb.workbookViews.length,
			externalReferenceCount: this.wb.externalReferences.length,
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
		const isHydrated = this.loadInfo.cellsHydrated
		const used = isHydrated ? sheet.cells.usedRange() : null
		const count = isHydrated ? sheet.cells.cellCount() : null
		const base = buildSheetInfo(sheet, isHydrated, used, count)
		const info = {
			...base,
			usedRange: used,
			state: sheet.state,
			merges: isHydrated ? [...sheet.merges] : null,
			tables: isHydrated ? sheet.tables.map((table) => buildTableInfo(table, sheet)) : null,
			comments: isHydrated
				? [...sheet.comments.entries()].map(([ref, comment]) => ({ ref, ...comment }))
				: null,
			hyperlinks: isHydrated
				? [...sheet.hyperlinks.entries()].map(([ref, hyperlink]) => ({ ref, ...hyperlink }))
				: null,
			ignoredErrors: isHydrated ? [...sheet.ignoredErrors] : null,
			conditionalFormats: isHydrated ? [...sheet.conditionalFormats] : null,
			dataValidations: isHydrated ? [...sheet.dataValidations] : null,
			imageRefs: isHydrated ? [...sheet.imageRefs] : null,
			drawingRefs: isHydrated ? { ...sheet.drawingRefs } : null,
			autoFilter: isHydrated ? sheet.autoFilter : null,
			protection: isHydrated ? sheet.protection : null,
			tabColor: isHydrated ? sheet.tabColor : null,
			sheetFormatPr: isHydrated ? sheet.sheetFormatPr : null,
			pageMargins: isHydrated ? sheet.pageMargins : null,
			pageSetup: isHydrated ? sheet.pageSetup : null,
			printOptions: isHydrated ? sheet.printOptions : null,
			headerFooter: isHydrated ? sheet.headerFooter : null,
		}
		this.sheetInspectCache.set(name, info)
		return info
	}

	sheet(name: string): SheetHandle | undefined {
		const cached = this.sheetHandleCache.get(name)
		if (cached) return cached
		const sheet = this.wb.getSheet(name)
		if (!sheet) return undefined
		const handle = new SheetHandle(sheet)
		this.sheetHandleCache.set(name, handle)
		return handle
	}

	readRange(sheetName: string, range: string): import('./types.ts').RangeInfo | undefined {
		return this.sheet(sheetName)?.range(range)
	}

	readWindow(
		sheetName: string,
		range: string,
		opts?: { rowOffset?: number; rowLimit?: number },
	): RangeWindowInfo | undefined {
		return this.sheet(sheetName)?.readWindow(range, opts)
	}

	*streamRange(
		sheetName: string,
		range: string,
	): Generator<readonly import('./types.ts').CellInfo[]> {
		const sheet = this.sheet(sheetName)
		if (!sheet) return
		yield* sheet.streamRange(range)
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

	table(name: string): TableHandle | undefined {
		const cached = this.tableHandleCache.get(name)
		if (cached) return cached
		for (const sheet of this.wb.sheets) {
			for (const table of sheet.tables) {
				if (table.name === name) {
					const handle = new TableHandle(table, sheet)
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
		const { sheetName, ref } = parseFullRef(cellRef, this.wb)
		const result = verifyTrace(this.wb, sheetName, ref, opts)
		if (!result.ok) return undefined
		return {
			ref: `${sheetName}!${ref}`,
			formula: result.value.formula,
			dependsOn: result.value.precedents.map((node) => `${node.sheet}!${node.ref}`),
			feedsInto: result.value.dependents.map((node) => `${node.sheet}!${node.ref}`),
		}
	}

	formula(cellRef: string): FormulaInfo | undefined {
		if (this.formulaInfoCache.has(cellRef)) return this.formulaInfoCache.get(cellRef)
		const { sheetName, ref } = parseFullRef(cellRef, this.wb)
		const cell = this.sheet(sheetName)?.cell(ref)
		if (!cell?.formula) {
			this.formulaInfoCache.set(cellRef, undefined)
			return undefined
		}

		const formula = normalizeFormulaInput(cell.formula)
		const tokens = tokenize(formula).filter(
			(token) => token.type !== 'Whitespace' && token.type !== 'EOF',
		)
		const parsed = parseFormula(formula)
		if (!parsed.ok) {
			const info = {
				ref: `${sheetName}!${ref}`,
				formula,
				normalizedFormula: formula,
				value: cell.value,
				...(cell.formulaBinding ? { binding: cell.formulaBinding } : {}),
				refs: [],
				functions: [],
				volatile: false,
				tokens,
				parseError: parsed.error.message,
			}
			this.formulaInfoCache.set(cellRef, info)
			return info
		}

		const ast = parsed.value
		const info = {
			ref: `${sheetName}!${ref}`,
			formula,
			normalizedFormula: printFormula(ast),
			value: cell.value,
			...(cell.formulaBinding ? { binding: cell.formulaBinding } : {}),
			refs: extractRefs(ast).map(formatFormulaRef),
			functions: [...collectFunctionNames(ast)],
			volatile: hasVolatileFunction(ast),
			tokens,
			ast,
		}
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
			name: entry.name,
			formula: entry.formula,
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
}

function buildSheetInfo(
	sheet: import('@ascend/core').Sheet,
	isHydrated: boolean,
	used: RangeRef | null,
	count: number | null,
): SheetInfo {
	return {
		name: sheet.name,
		rowCount: used ? used.end.row + 1 : null,
		colCount: used ? used.end.col + 1 : null,
		cellCount: count,
		tableCount: isHydrated ? sheet.tables.length : null,
		commentCount: isHydrated ? sheet.comments.size : null,
		conditionalFormatCount: isHydrated ? sheet.conditionalFormats.length : null,
		dataValidationCount: isHydrated ? sheet.dataValidations.length : null,
		hasFrozenPanes: isHydrated ? sheet.frozenRows > 0 || sheet.frozenCols > 0 : null,
		colWidthCount: isHydrated ? sheet.colWidths.size : null,
		imageCount: isHydrated ? sheet.imageRefs.length : null,
		rowHeightCount: isHydrated ? sheet.rowHeights.size : null,
		hyperlinkCount: isHydrated ? sheet.hyperlinks.size : null,
		ignoredErrorCount: isHydrated ? sheet.ignoredErrors.length : null,
		hasAutoFilter: isHydrated ? sheet.autoFilter !== null : null,
		hasDrawingRefs: isHydrated
			? sheet.drawingRefs.hasDrawing || sheet.drawingRefs.hasLegacyDrawing
			: null,
		hasPageMetadata: isHydrated
			? sheet.pageMargins !== null ||
				sheet.pageSetup !== null ||
				sheet.printOptions !== null ||
				sheet.headerFooter !== null
			: null,
		hasProtection: isHydrated ? sheet.protection !== null : null,
		cellDataLoaded: isHydrated,
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

function formatFormulaRef(ref: import('@ascend/formulas').FormulaRef): string {
	if (ref.kind === 'cell') {
		return `${ref.sheet ? `${ref.sheet}!` : ''}${formatFormulaCellRef(ref.ref)}`
	}
	if (ref.kind === 'range') {
		return `${ref.sheet ? `${ref.sheet}!` : ''}${formatFormulaCellRef(ref.start)}:${formatFormulaCellRef(ref.end)}`
	}
	if (ref.kind === 'wholeRowRange') {
		return `${ref.sheet ? `${ref.sheet}!` : ''}${ref.startRow + 1}:${ref.endRow + 1}`
	}
	return `${ref.sheet ? `${ref.sheet}!` : ''}${indexToColumn(ref.startCol)}:${indexToColumn(ref.endCol)}`
}

function formatFormulaCellRef(ref: FormulaCellRef): string {
	return `${ref.colAbsolute ? '$' : ''}${indexToColumn(ref.col)}${ref.rowAbsolute ? '$' : ''}${ref.row + 1}`
}

function hasVolatileFunction(node: FormulaNode): boolean {
	switch (node.type) {
		case 'function':
			if (functionRegistry.get(node.name.toUpperCase())?.volatile) return true
			return node.args.some(hasVolatileFunction)
		case 'binary':
			return hasVolatileFunction(node.left) || hasVolatileFunction(node.right)
		case 'unary':
			return hasVolatileFunction(node.operand)
		case 'array':
			return node.rows.some((row) => row.some(hasVolatileFunction))
		case 'spillRef':
			return hasVolatileFunction(node.target)
		case 'wholeRowRange':
		case 'wholeColumnRange':
			return false
		default:
			return false
	}
}

function collectFunctionNames(node: FormulaNode, out = new Set<string>()): Set<string> {
	switch (node.type) {
		case 'function':
			out.add(node.name)
			for (const arg of node.args) collectFunctionNames(arg, out)
			break
		case 'binary':
			collectFunctionNames(node.left, out)
			collectFunctionNames(node.right, out)
			break
		case 'unary':
			collectFunctionNames(node.operand, out)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) collectFunctionNames(cell, out)
			}
			break
		case 'spillRef':
			collectFunctionNames(node.target, out)
			break
		case 'wholeRowRange':
		case 'wholeColumnRange':
			break
	}
	return out
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
