import { readFile, writeFile } from 'node:fs/promises'
import {
	createWorkbook,
	indexToColumn,
	parseRange,
	type RangeRef,
	type Workbook,
} from '@ascend/core'
import {
	applyOperations,
	type CalcContext,
	createSnapshot,
	defaultCalcContext,
	diffWorkbooks,
	recalculate,
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
import { readCsv, writeCsv } from '@ascend/io-csv'
import {
	type PreservationCapsule,
	type ReadXlsxLoadInfo,
	readXlsx,
	writeXlsx,
} from '@ascend/io-xlsx'
import {
	type CompatibilityReport,
	type CsvDialect,
	emptyReport,
	type Operation,
} from '@ascend/schema'
import { check as verifyCheck, lint as verifyLint, trace as verifyTrace } from '@ascend/verify'
import { SheetHandle } from './sheet-handle.ts'
import { TableHandle } from './table-handle.ts'
import type {
	ApplyResult,
	CheckIssue,
	CheckResult,
	FormulaInfo,
	LintResult,
	LintWarning,
	RangeWindowInfo,
	RecalcResult,
	SheetInfo,
	SheetInspectInfo,
	TableInfo,
	TraceResult,
	WorkbookInfo,
} from './types.ts'

const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04]

function isZip(bytes: Uint8Array): boolean {
	if (bytes.length < 4) return false
	return (
		bytes[0] === ZIP_MAGIC[0] &&
		bytes[1] === ZIP_MAGIC[1] &&
		bytes[2] === ZIP_MAGIC[2] &&
		bytes[3] === ZIP_MAGIC[3]
	)
}

function cloneWorkbook(source: Workbook): Workbook {
	return source.clone()
}

export class AscendWorkbook {
	private readonly wb: Workbook
	private readonly caps: PreservationCapsule[]
	private readonly compat: CompatibilityReport
	private readonly loadInfo: import('./types.ts').WorkbookLoadInfo
	private originalBytes: Uint8Array | null
	private dirty: boolean
	private readonly dirtySheets = new Set<string>()
	private workbookMetaDirty = false
	private sharedStringsDirty = false

	private constructor(
		workbook: Workbook,
		capsules: PreservationCapsule[],
		report: CompatibilityReport,
		loadInfo: import('./types.ts').WorkbookLoadInfo,
		originalBytes: Uint8Array | null,
	) {
		this.wb = workbook
		this.caps = capsules
		this.compat = report
		this.loadInfo = loadInfo
		this.originalBytes = originalBytes
		this.dirty = false
	}

	static async open(
		pathOrBytes: string | Uint8Array,
		options?: { mode?: 'full' | 'metadata-only' | 'values'; sheets?: readonly string[] },
	): Promise<AscendWorkbook> {
		let bytes: Uint8Array
		let ext = ''

		if (typeof pathOrBytes === 'string') {
			ext = pathOrBytes.split('.').pop()?.toLowerCase() ?? ''
			bytes = new Uint8Array(await readFile(pathOrBytes))
		} else {
			bytes = pathOrBytes
		}

		if (ext === 'csv' || ext === 'tsv') {
			const text = new TextDecoder().decode(bytes)
			const dialect: Partial<CsvDialect> | undefined =
				ext === 'tsv' ? { delimiter: '\t' } : undefined
			const result = readCsv(text, dialect)
			if (!result.ok) throw new Error(result.error.message)
			return new AscendWorkbook(
				result.value,
				[],
				emptyReport('csv'),
				buildLoadInfo({
					mode: 'full',
					isPartial: false,
					cellsHydrated: true,
					hasAllSheets: true,
					sourceSheetNames: result.value.sheets.map((sheet) => sheet.name),
					loadedSheetNames: result.value.sheets.map((sheet) => sheet.name),
				}),
				null,
			)
		}

		if (ext === 'xlsx' || ext === 'xlsm' || isZip(bytes)) {
			const result = readXlsx(bytes, options)
			if (!result.ok) throw new Error(result.error.message)
			const loadInfo = buildLoadInfo(result.value.loadInfo)
			if (loadInfo.isPartial) {
				result.value.workbook.sourceArchiveBytes = null
			}
			return new AscendWorkbook(
				result.value.workbook,
				result.value.capsules,
				result.value.report,
				loadInfo,
				loadInfo.isPartial ? null : bytes,
			)
		}

		const text = new TextDecoder().decode(bytes)
		const result = readCsv(text)
		if (!result.ok) throw new Error(result.error.message)
		return new AscendWorkbook(
			result.value,
			[],
			emptyReport('csv'),
			buildLoadInfo({
				mode: 'full',
				isPartial: false,
				cellsHydrated: true,
				hasAllSheets: true,
				sourceSheetNames: result.value.sheets.map((sheet) => sheet.name),
				loadedSheetNames: result.value.sheets.map((sheet) => sheet.name),
			}),
			null,
		)
	}

	static create(): AscendWorkbook {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		return new AscendWorkbook(
			wb,
			[],
			emptyReport('ascend'),
			buildLoadInfo({
				mode: 'full',
				isPartial: false,
				cellsHydrated: true,
				hasAllSheets: true,
				sourceSheetNames: ['Sheet1'],
				loadedSheetNames: ['Sheet1'],
			}),
			null,
		)
	}

	static fromCsv(content: string, dialect?: Partial<CsvDialect>): AscendWorkbook {
		const result = readCsv(content, dialect)
		if (!result.ok) throw new Error(result.error.message)
		return new AscendWorkbook(
			result.value,
			[],
			emptyReport('csv'),
			buildLoadInfo({
				mode: 'full',
				isPartial: false,
				cellsHydrated: true,
				hasAllSheets: true,
				sourceSheetNames: result.value.sheets.map((sheet) => sheet.name),
				loadedSheetNames: result.value.sheets.map((sheet) => sheet.name),
			}),
			null,
		)
	}

	// --- Inspection ---

	inspect(): WorkbookInfo {
		let totalCells = 0
		let totalComments = 0
		let totalConditionalFormats = 0
		let totalDataValidations = 0
		let totalImages = 0
		const sheets = this.wb.sheets.map((s) => {
			const isHydrated = this.loadInfo.cellsHydrated
			const used = isHydrated ? s.cells.usedRange() : null
			const count = isHydrated ? s.cells.cellCount() : null
			if (count !== null) totalCells += count
			if (isHydrated) {
				totalComments += s.comments.size
				totalConditionalFormats += s.conditionalFormats.length
				totalDataValidations += s.dataValidations.length
				totalImages += s.imageRefs.length
			}
			return buildSheetInfo(s, isHydrated, used, count)
		})
		return {
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
	}

	inspectSheet(name: string): SheetInspectInfo | undefined {
		const sheet = this.wb.getSheet(name)
		if (!sheet) return undefined
		const isHydrated = this.loadInfo.cellsHydrated
		const used = isHydrated ? sheet.cells.usedRange() : null
		const count = isHydrated ? sheet.cells.cellCount() : null
		const base = buildSheetInfo(sheet, isHydrated, used, count)
		return {
			...base,
			usedRange: used,
			state: sheet.state,
			merges: isHydrated ? [...sheet.merges] : null,
			tables: isHydrated ? sheet.tables.map((table) => buildTableInfo(table)) : null,
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
	}

	sheet(name: string): SheetHandle | undefined {
		const s = this.wb.getSheet(name)
		return s ? new SheetHandle(s) : undefined
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
		for (const sheet of this.wb.sheets) {
			for (const tbl of sheet.tables) {
				if (tbl.name === name) return new TableHandle(tbl, sheet)
			}
		}
		return undefined
	}

	// --- Mutation ---

	preview(ops: readonly Operation[]): import('./types.ts').PreviewResult {
		const clone = cloneWorkbook(this.wb)
		const errors: import('@ascend/schema').AscendError[] = []

		const result = applyOperations(clone, ops)
		if (!result.ok) {
			errors.push(result.error)
			return {
				diff: { sheets: [], namesAdded: [], namesRemoved: [], namesChanged: [] },
				sheetDiffs: [],
				cellChanges: [],
				errors,
			}
		}

		if (result.value.recalcRequired) {
			recalculate(
				clone,
				defaultCalcContext({
					dateSystem: clone.calcSettings.dateSystem,
					iterativeCalc: clone.calcSettings.iterativeCalc,
				}),
			)
		}

		const diff = diffWorkbooks(this.wb, clone)
		const cellChanges = diff.sheets.flatMap((s) => s.cellsChanged)

		return { diff, sheetDiffs: diff.sheets, cellChanges, errors }
	}

	apply(ops: readonly Operation[]): ApplyResult {
		const dirtyFlags = this.deriveDirtyFlags(ops)
		const result = applyOperations(this.wb, ops)
		if (!result.ok) {
			return {
				affectedCells: [],
				sheetsModified: [],
				recalcRequired: false,
				errors: [result.error],
			}
		}

		this.markDirty()
		for (const sheetName of result.value.sheetsModified) this.dirtySheets.add(sheetName)
		this.workbookMetaDirty ||= dirtyFlags.workbookMetaDirty
		this.sharedStringsDirty ||= dirtyFlags.sharedStringsDirty
		return {
			affectedCells: result.value.affectedCells,
			sheetsModified: result.value.sheetsModified,
			recalcRequired: result.value.recalcRequired,
			errors: [],
		}
	}

	recalc(opts?: { range?: string }): RecalcResult {
		const ctx: CalcContext = defaultCalcContext({
			dateSystem: this.wb.calcSettings.dateSystem,
			iterativeCalc: this.wb.calcSettings.iterativeCalc,
		})
		let rangeRef: RangeRef | undefined
		if (opts?.range) {
			rangeRef = parseRange(opts.range)
		}
		const result = recalculate(this.wb, ctx, rangeRef ? { range: rangeRef } : undefined)
		if (result.changed.length > 0 || result.errors.length > 0) {
			this.markDirty()
			this.sharedStringsDirty = true
			for (const ref of result.changed) {
				const bang = ref.indexOf('!')
				if (bang !== -1) this.dirtySheets.add(ref.slice(0, bang))
			}
		}
		return {
			changed: result.changed,
			errors: result.errors,
			duration: result.duration,
		}
	}

	// --- Verification ---

	check(): CheckResult {
		const result = verifyCheck(this.wb)
		const issues: CheckIssue[] = result.issues.map((issue) =>
			issue.refs?.[0]
				? {
						severity: issue.severity === 'info' ? 'warning' : issue.severity,
						message: issue.message,
						ref: issue.refs[0],
					}
				: {
						severity: issue.severity === 'info' ? 'warning' : issue.severity,
						message: issue.message,
					},
		)
		return { valid: result.passed, issues }
	}

	lint(): LintResult {
		const result = verifyLint(this.wb)
		const warnings: LintWarning[] = result.violations.map((violation) => ({
			rule: violation.rule,
			message: violation.message,
			ref: violation.ref,
		}))
		return { clean: warnings.length === 0, warnings }
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
		const { sheetName, ref } = parseFullRef(cellRef, this.wb)
		const cell = this.sheet(sheetName)?.cell(ref)
		if (!cell?.formula) return undefined

		const formula = normalizeFormulaInput(cell.formula)
		const tokens = tokenize(formula).filter(
			(token) => token.type !== 'Whitespace' && token.type !== 'EOF',
		)
		const parsed = parseFormula(formula)
		if (!parsed.ok) {
			return {
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
		}

		const ast = parsed.value
		return {
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
	}

	setFormula(cellRef: string, formula: string): ApplyResult {
		const { sheetName, ref } = parseFullRef(cellRef, this.wb)
		return this.apply([
			{ op: 'setFormula', sheet: sheetName, ref, formula: normalizeFormulaInput(formula) },
		])
	}

	fillFormula(rangeRef: string, formula: string): ApplyResult {
		const { sheetName, ref } = parseFullRef(rangeRef, this.wb)
		return this.apply([
			{ op: 'fillFormula', sheet: sheetName, range: ref, formula: normalizeFormulaInput(formula) },
		])
	}

	diff(other: AscendWorkbook): WorkbookDiff {
		return diffWorkbooks(this.wb, other.wb)
	}

	snapshot(): WorkbookSnapshot {
		return createSnapshot(this.wb)
	}

	// --- Export ---

	async save(path: string): Promise<void> {
		this.assertWritable()
		const ext = path.split('.').pop()?.toLowerCase() ?? ''

		if (ext === 'csv' || ext === 'tsv') {
			const result =
				ext === 'tsv' ? writeCsv(this.wb, { dialect: { delimiter: '\t' } }) : writeCsv(this.wb)
			if (!result.ok) throw new Error(result.error.message)
			await writeFile(path, result.value, 'utf-8')
			return
		}

		const bytes = this.toBytes()
		await writeFile(path, bytes)
	}

	toBytes(): Uint8Array {
		this.assertWritable()
		if (this.originalBytes && !this.dirty) return this.originalBytes
		const result = writeXlsx(this.wb, this.caps.length > 0 ? this.caps : undefined, {
			dirtySheetNames: [...this.dirtySheets],
			workbookMetaDirty: this.workbookMetaDirty,
			sharedStringsDirty: this.sharedStringsDirty,
		})
		if (!result.ok) throw new Error(result.error.message)
		this.captureSerializedState(result.value)
		return result.value
	}

	toCsv(opts?: { sheet?: string; range?: string }): string {
		this.assertWritable()
		const result = writeCsv(this.wb, opts)
		if (!result.ok) throw new Error(result.error.message)
		return result.value
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

	// --- Access ---

	get report(): CompatibilityReport {
		return this.compat
	}

	get sheets(): readonly string[] {
		return this.wb.sheets.map((s) => s.name)
	}

	get names(): readonly string[] {
		return this.wb.definedNames.workbookKeys()
	}

	definedName(
		name: string,
		scopeSheetName?: string,
	): import('./types.ts').DefinedNameInfo | undefined {
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

	private assertWritable(): void {
		if (!this.loadInfo.isPartial) return
		throw new Error(
			'Cannot export a partial workbook view. Reopen the workbook with a full load before saving or exporting.',
		)
	}

	private deriveDirtyFlags(ops: readonly Operation[]): {
		workbookMetaDirty: boolean
		sharedStringsDirty: boolean
	} {
		let workbookMetaDirty = false
		let sharedStringsDirty = false
		let sharedStringKeys: Set<string> | null = null
		const getSharedStringKeys = (): Set<string> => {
			if (sharedStringKeys) return sharedStringKeys
			sharedStringKeys = collectSharedStringKeys(this.wb)
			return sharedStringKeys
		}
		for (const op of ops) {
			switch (op.op) {
				case 'addSheet':
				case 'deleteSheet':
				case 'renameSheet':
				case 'moveSheet':
				case 'setDefinedName':
				case 'deleteDefinedName':
					workbookMetaDirty = true
					break
				case 'setFormula':
				case 'fillFormula':
					sharedStringsDirty = true
					break
				case 'setCells':
					if (
						op.updates.some((update) => {
							if (typeof update.value !== 'string') return false
							return !getSharedStringKeys().has(makePlainSharedStringKey(update.value))
						})
					) {
						sharedStringsDirty = true
					}
					break
				case 'appendRows':
					if (
						op.rows.some((row) =>
							row.some(
								(value) =>
									typeof value === 'string' &&
									!getSharedStringKeys().has(makePlainSharedStringKey(value)),
							),
						)
					) {
						sharedStringsDirty = true
					}
					break
			}
		}
		return { workbookMetaDirty, sharedStringsDirty }
	}

	private markDirty(): void {
		if (!this.dirty) this.originalBytes = null
		this.dirty = true
	}

	private captureSerializedState(bytes: Uint8Array): void {
		this.originalBytes = bytes
		this.wb.sourceArchiveBytes = bytes
		this.dirty = false
		this.dirtySheets.clear()
		this.workbookMetaDirty = false
		this.sharedStringsDirty = false
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

function buildTableInfo(table: import('@ascend/core').Table): TableInfo {
	const headerOffset = table.hasHeaders ? 1 : 0
	const totalOffset = table.hasTotals ? 1 : 0
	return {
		name: table.name,
		ref: table.ref,
		rowCount: table.ref.end.row - table.ref.start.row + 1 - headerOffset - totalOffset,
		hasHeaders: table.hasHeaders,
		hasTotals: table.hasTotals,
		autoFilter: table.autoFilter ?? null,
		...(table.sortState?.ref ? { sortStateRef: table.sortState.ref } : {}),
		...(table.tableStyleInfo ? { styleInfo: table.tableStyleInfo } : {}),
		columnDefs: [...table.columns],
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

function buildLoadInfo(info: ReadXlsxLoadInfo): import('./types.ts').WorkbookLoadInfo {
	return {
		mode: info.mode,
		isPartial: info.isPartial,
		cellsHydrated: info.cellsHydrated,
		hasAllSheets: info.hasAllSheets,
		sourceSheets: info.sourceSheetNames,
		loadedSheets: info.loadedSheetNames,
	}
}

function collectSharedStringKeys(workbook: Workbook): Set<string> {
	const keys = new Set<string>()
	for (const sheet of workbook.sheets) {
		for (const [, , cell] of sheet.cells.iterate()) {
			const key = makeSharedStringKey(cell.value)
			if (key) keys.add(key)
		}
	}
	return keys
}

function makePlainSharedStringKey(value: string): string {
	return `s:${value}`
}

function makeSharedStringKey(value: import('@ascend/schema').CellValue | string): string | null {
	if (typeof value === 'string') return makePlainSharedStringKey(value)
	if (value.kind === 'string') return `s:${value.value}`
	if (value.kind === 'richText') return `r:${JSON.stringify(value.runs)}`
	return null
}

function normalizeFormulaInput(formula: string): string {
	return formula.startsWith('=') ? formula.slice(1) : formula
}

function formatFormulaRef(ref: import('@ascend/formulas').FormulaRef): string {
	if (ref.kind === 'cell') {
		return `${ref.sheet ? `${ref.sheet}!` : ''}${formatFormulaCellRef(ref.ref)}`
	}
	return `${ref.sheet ? `${ref.sheet}!` : ''}${formatFormulaCellRef(ref.start)}:${formatFormulaCellRef(ref.end)}`
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
