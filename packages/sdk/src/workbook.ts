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
	const clone = createWorkbook()
	clone.calcSettings = source.calcSettings
	clone.workbookProperties = { ...source.workbookProperties }
	clone.workbookProtection = source.workbookProtection ? { ...source.workbookProtection } : null
	clone.styleMetadata = { ...source.styleMetadata }
	clone.themeMetadata = { ...source.themeMetadata }
	clone.preservedStyles = source.preservedStyles
		? {
				xml: source.preservedStyles.xml,
				xfByStyleId: { ...source.preservedStyles.xfByStyleId },
			}
		: null
	clone.preservedTheme = source.preservedTheme
		? {
				path: source.preservedTheme.path,
				contentType: source.preservedTheme.contentType,
				xml: source.preservedTheme.xml,
			}
		: null
	clone.preservedXml = source.preservedXml
		? {
				workbookXml: source.preservedXml.workbookXml,
				...(source.preservedXml.workbookRelsXml
					? { workbookRelsXml: source.preservedXml.workbookRelsXml }
					: {}),
			}
		: null
	clone.workbookViews.push(...source.workbookViews.map((view) => ({ ...view })))
	clone.externalReferences.push(...source.externalReferences)
	clone.differentialStyles.push(...source.differentialStyles.map((style) => ({ ...style })))

	for (const definedName of source.definedNames.list()) {
		clone.definedNames.set(definedName.name, definedName.formula, definedName.scope)
	}

	for (const sheet of source.sheets) {
		const cloned = clone.addSheet(sheet.name)
		for (const [row, col, cell] of sheet.cells.iterate()) {
			cloned.cells.set(row, col, { ...cell })
		}
		for (const merge of sheet.merges) {
			cloned.merges.push(merge)
		}
		for (const table of sheet.tables) {
			cloned.tables.push(table)
		}
		for (const colDef of sheet.colDefs) {
			cloned.colDefs.push({ ...colDef })
		}
		cloned.state = sheet.state
		for (const [k, v] of sheet.colWidths) cloned.colWidths.set(k, v)
		for (const [k, v] of sheet.rowHeights) cloned.rowHeights.set(k, v)
		cloned.frozenRows = sheet.frozenRows
		cloned.frozenCols = sheet.frozenCols
		for (const [k, v] of sheet.comments) cloned.comments.set(k, v)
		for (const [k, v] of sheet.hyperlinks) cloned.hyperlinks.set(k, { ...v })
		cloned.ignoredErrors.push(...sheet.ignoredErrors)
		cloned.dataValidations.push(...sheet.dataValidations.map((validation) => ({ ...validation })))
		cloned.conditionalFormats.push(
			...sheet.conditionalFormats.map((conditionalFormat) => ({
				sqref: conditionalFormat.sqref,
				rules: conditionalFormat.rules.map((rule) => ({
					...rule,
					formulas: [...rule.formulas],
					...(rule.style ? { style: { ...rule.style } } : {}),
				})),
			})),
		)
		cloned.drawingRefs = { ...sheet.drawingRefs }
		cloned.autoFilter = sheet.autoFilter
		cloned.protection = sheet.protection ? { ...sheet.protection } : null
		cloned.pageMargins = sheet.pageMargins ? { ...sheet.pageMargins } : null
		cloned.pageSetup = sheet.pageSetup ? { ...sheet.pageSetup } : null
		cloned.printOptions = sheet.printOptions ? { ...sheet.printOptions } : null
		cloned.headerFooter = sheet.headerFooter ? { ...sheet.headerFooter } : null
		cloned.preservedXml = sheet.preservedXml
			? {
					xml: sheet.preservedXml.xml,
					...(sheet.preservedXml.relsXml ? { relsXml: sheet.preservedXml.relsXml } : {}),
				}
			: null
	}

	return clone
}

export class AscendWorkbook {
	private readonly wb: Workbook
	private readonly caps: PreservationCapsule[]
	private readonly compat: CompatibilityReport
	private readonly loadInfo: import('./types.ts').WorkbookLoadInfo
	private readonly originalBytes: Uint8Array | null
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
		options?: { mode?: 'full' | 'metadata-only'; sheets?: readonly string[] },
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
			return new AscendWorkbook(
				result.value.workbook,
				result.value.capsules,
				result.value.report,
				buildLoadInfo(result.value.loadInfo),
				bytes,
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
		const sheets = this.wb.sheets.map((s) => {
			const isHydrated = this.loadInfo.cellsHydrated
			const used = isHydrated ? s.cells.usedRange() : null
			const count = isHydrated ? s.cells.cellCount() : null
			if (count !== null) totalCells += count
			const info: import('./types.ts').SheetInfo = {
				name: s.name,
				rowCount: used ? used.end.row + 1 : null,
				colCount: used ? used.end.col + 1 : null,
				cellCount: count,
				tableCount: isHydrated ? s.tables.length : null,
				hasFrozenPanes: isHydrated ? s.frozenRows > 0 || s.frozenCols > 0 : null,
				colWidthCount: isHydrated ? s.colWidths.size : null,
				rowHeightCount: isHydrated ? s.rowHeights.size : null,
				hyperlinkCount: isHydrated ? s.hyperlinks.size : null,
				ignoredErrorCount: isHydrated ? s.ignoredErrors.length : null,
				hasAutoFilter: isHydrated ? s.autoFilter !== null : null,
				hasPageMetadata: isHydrated
					? s.pageMargins !== null ||
						s.pageSetup !== null ||
						s.printOptions !== null ||
						s.headerFooter !== null
					: null,
				cellDataLoaded: isHydrated,
			}
			return info
		})
		return {
			sheetCount: this.loadInfo.sourceSheets.length,
			loadedSheetCount: this.loadInfo.loadedSheets.length,
			sheets,
			definedNames: this.wb.definedNames.workbookKeys(),
			cellCount: this.loadInfo.cellsHydrated ? totalCells : null,
			sourceFormat: this.compat.sourceFormat,
			workbookViewCount: this.wb.workbookViews.length,
			externalReferenceCount: this.wb.externalReferences.length,
			styleSummary: { ...this.wb.styleMetadata },
			themeSummary: {
				hasThemePart: this.wb.preservedTheme !== null,
				...this.wb.themeMetadata,
			},
			compatibility: this.compat,
			load: this.loadInfo,
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
		const result = applyOperations(this.wb, ops)
		if (!result.ok) {
			return {
				affectedCells: [],
				sheetsModified: [],
				recalcRequired: false,
				errors: [result.error],
			}
		}

		this.dirty = true
		for (const sheetName of result.value.sheetsModified) this.dirtySheets.add(sheetName)
		this.updateDirtyFlags(ops)
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
			this.dirty = true
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

	private updateDirtyFlags(ops: readonly Operation[]): void {
		for (const op of ops) {
			switch (op.op) {
				case 'addSheet':
				case 'deleteSheet':
				case 'renameSheet':
				case 'moveSheet':
				case 'setDefinedName':
				case 'deleteDefinedName':
					this.workbookMetaDirty = true
					break
				case 'setFormula':
				case 'clearRange':
					this.sharedStringsDirty = true
					break
				case 'setCells':
					if (op.updates.some((update) => typeof update.value === 'string')) {
						this.sharedStringsDirty = true
					}
					break
			}
		}
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
