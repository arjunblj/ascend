import { readFile, writeFile } from 'node:fs/promises'
import { createWorkbook, parseRange, type RangeRef, type Workbook } from '@ascend/core'
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
import { readCsv, writeCsv } from '@ascend/io-csv'
import { type PreservationCapsule, readXlsx, writeXlsx } from '@ascend/io-xlsx'
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
		cloned.state = sheet.state
		for (const [k, v] of sheet.colWidths) cloned.colWidths.set(k, v)
		for (const [k, v] of sheet.rowHeights) cloned.rowHeights.set(k, v)
		cloned.frozenRows = sheet.frozenRows
		cloned.frozenCols = sheet.frozenCols
		for (const [k, v] of sheet.comments) cloned.comments.set(k, v)
	}

	return clone
}

export class AscendWorkbook {
	private readonly wb: Workbook
	private readonly caps: PreservationCapsule[]
	private readonly compat: CompatibilityReport

	private constructor(
		workbook: Workbook,
		capsules: PreservationCapsule[],
		report: CompatibilityReport,
	) {
		this.wb = workbook
		this.caps = capsules
		this.compat = report
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
			return new AscendWorkbook(result.value, [], emptyReport('csv'))
		}

		if (ext === 'xlsx' || ext === 'xlsm' || isZip(bytes)) {
			const result = readXlsx(bytes, options)
			if (!result.ok) throw new Error(result.error.message)
			return new AscendWorkbook(result.value.workbook, result.value.capsules, result.value.report)
		}

		const text = new TextDecoder().decode(bytes)
		const result = readCsv(text)
		if (!result.ok) throw new Error(result.error.message)
		return new AscendWorkbook(result.value, [], emptyReport('csv'))
	}

	static create(): AscendWorkbook {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		return new AscendWorkbook(wb, [], emptyReport('ascend'))
	}

	static fromCsv(content: string, dialect?: Partial<CsvDialect>): AscendWorkbook {
		const result = readCsv(content, dialect)
		if (!result.ok) throw new Error(result.error.message)
		return new AscendWorkbook(result.value, [], emptyReport('csv'))
	}

	// --- Inspection ---

	inspect(): WorkbookInfo {
		let totalCells = 0
		const sheets = this.wb.sheets.map((s) => {
			const used = s.cells.usedRange()
			const count = s.cells.cellCount()
			totalCells += count
			const info: import('./types.ts').SheetInfo = {
				name: s.name,
				rowCount: used ? used.end.row + 1 : 0,
				colCount: used ? used.end.col + 1 : 0,
				cellCount: count,
				tableCount: s.tables.length,
				hasFrozenPanes: s.frozenRows > 0 || s.frozenCols > 0,
			}
			return info
		})
		return {
			sheetCount: this.wb.sheets.length,
			sheets,
			definedNames: this.wb.definedNames.workbookKeys(),
			cellCount: totalCells,
			sourceFormat: this.compat.sourceFormat,
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

	trace(cellRef: string): TraceResult | undefined {
		const { sheetName, ref } = parseFullRef(cellRef, this.wb)
		const result = verifyTrace(this.wb, sheetName, ref)
		if (!result.ok) return undefined
		return {
			ref: `${sheetName}!${ref}`,
			formula: result.value.formula,
			dependsOn: result.value.precedents.map((node) => `${node.sheet}!${node.ref}`),
			feedsInto: result.value.dependents.map((node) => `${node.sheet}!${node.ref}`),
		}
	}

	diff(other: AscendWorkbook): WorkbookDiff {
		return diffWorkbooks(this.wb, other.wb)
	}

	snapshot(): WorkbookSnapshot {
		return createSnapshot(this.wb)
	}

	// --- Export ---

	async save(path: string): Promise<void> {
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
		const result = writeXlsx(this.wb, this.caps.length > 0 ? this.caps : undefined)
		if (!result.ok) throw new Error(result.error.message)
		return result.value
	}

	toCsv(opts?: { sheet?: string; range?: string }): string {
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
