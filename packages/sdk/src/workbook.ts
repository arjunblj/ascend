import { readFile, writeFile } from 'node:fs/promises'
import { createWorkbook, parseRange, type RangeRef, type Workbook } from '@ascend/core'
import {
	applyOperations,
	type CalcContext,
	cellKey,
	createSnapshot,
	defaultCalcContext,
	diffWorkbooks,
	recalculate,
	type WorkbookDiff,
	type WorkbookSnapshot,
} from '@ascend/engine'
import { extractRefs, parseFormula } from '@ascend/formulas'
import { readCsv, writeCsv } from '@ascend/io-csv'
import { type PreservationCapsule, readXlsx, writeXlsx } from '@ascend/io-xlsx'
import {
	type CompatibilityReport,
	type CsvDialect,
	emptyReport,
	type Operation,
} from '@ascend/schema'
import { SheetHandle } from './sheet-handle.ts'
import { TableHandle } from './table-handle.ts'
import type {
	ApplyResult,
	CheckIssue,
	CheckResult,
	LintResult,
	LintWarning,
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

	for (const [name, ref] of source.definedNames) {
		clone.definedNames.set(name, ref)
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

	static async open(pathOrBytes: string | Uint8Array): Promise<AscendWorkbook> {
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
			const result = readXlsx(bytes)
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
			definedNames: [...this.wb.definedNames.keys()],
			cellCount: totalCells,
			sourceFormat: this.compat.sourceFormat,
		}
	}

	sheet(name: string): SheetHandle | undefined {
		const s = this.wb.getSheet(name)
		return s ? new SheetHandle(s) : undefined
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
			recalculate(clone, defaultCalcContext())
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
		const ctx: CalcContext = defaultCalcContext()
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
		const issues: CheckIssue[] = []
		const sheetNames = new Set<string>()

		for (const sheet of this.wb.sheets) {
			if (sheetNames.has(sheet.name)) {
				issues.push({ severity: 'error', message: `Duplicate sheet name: "${sheet.name}"` })
			}
			sheetNames.add(sheet.name)

			for (const [row, col, cell] of sheet.cells.iterate()) {
				if (cell.formula) {
					const parsed = parseFormula(cell.formula)
					if (!parsed.ok) {
						const ref = `${sheet.name}!${colLabel(col)}${row + 1}`
						issues.push({
							severity: 'error',
							message: `Invalid formula in ${ref}: ${cell.formula}`,
							ref,
						})
					}
				}
			}
		}

		if (this.wb.sheets.length === 0) {
			issues.push({ severity: 'warning', message: 'Workbook has no sheets' })
		}

		return { valid: issues.filter((i) => i.severity === 'error').length === 0, issues }
	}

	lint(): LintResult {
		const warnings: LintWarning[] = []
		const volatileFns = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT'])

		for (const sheet of this.wb.sheets) {
			for (const [row, col, cell] of sheet.cells.iterate()) {
				if (!cell.formula) continue
				const ref = `${sheet.name}!${colLabel(col)}${row + 1}`

				const parsed = parseFormula(cell.formula)
				if (!parsed.ok) {
					warnings.push({
						rule: 'parse-error',
						message: `Unparseable formula: ${cell.formula}`,
						ref,
					})
					continue
				}

				const volatiles = findVolatileCalls(parsed.value, volatileFns)
				for (const fn of volatiles) {
					warnings.push({
						rule: 'volatile-function',
						message: `Uses volatile function ${fn}()`,
						ref,
					})
				}

				if (cell.value.kind === 'error') {
					warnings.push({
						rule: 'error-value',
						message: `Cell contains error: ${cell.value.value}`,
						ref,
					})
				}
			}
		}

		return { clean: warnings.length === 0, warnings }
	}

	trace(cellRef: string): TraceResult | undefined {
		const { sheetName, row, col } = parseFullRef(cellRef, this.wb)
		const sheet = this.wb.getSheet(sheetName)
		if (!sheet) return undefined

		const cell = sheet.cells.get(row, col)
		const dependsOn: string[] = []
		const feedsInto: string[] = []

		if (cell?.formula) {
			const parsed = parseFormula(cell.formula)
			if (parsed.ok) {
				const refs = extractRefs(parsed.value)
				for (const ref of refs) {
					if (ref.kind === 'cell') {
						const refSheet = ref.sheet ?? sheetName
						dependsOn.push(`${refSheet}!${colLabel(ref.ref.col)}${ref.ref.row + 1}`)
					} else {
						const refSheet = ref.sheet ?? sheetName
						const start = `${colLabel(ref.start.col)}${ref.start.row + 1}`
						const end = `${colLabel(ref.end.col)}${ref.end.row + 1}`
						dependsOn.push(`${refSheet}!${start}:${end}`)
					}
				}
			}
		}

		const sheetIndex = this.wb.sheets.findIndex((s) => s.name === sheetName)
		if (sheetIndex >= 0) {
			const targetKey = cellKey(sheetIndex, row, col)
			for (let si = 0; si < this.wb.sheets.length; si++) {
				const s = this.wb.sheets[si]
				if (!s) continue
				for (const [r, c, otherCell] of s.cells.iterate()) {
					if (!otherCell.formula) continue
					const parsed = parseFormula(otherCell.formula)
					if (!parsed.ok) continue
					const refs = extractRefs(parsed.value)
					for (const ref of refs) {
						if (ref.kind === 'cell') {
							const refSi = ref.sheet
								? this.wb.sheets.findIndex(
										(ws) => ws.name.toLowerCase() === ref.sheet?.toLowerCase(),
									)
								: si
							if (refSi >= 0 && cellKey(refSi, ref.ref.row, ref.ref.col) === targetKey) {
								feedsInto.push(`${s.name}!${colLabel(c)}${r + 1}`)
							}
						} else {
							const refSi = ref.sheet
								? this.wb.sheets.findIndex(
										(ws) => ws.name.toLowerCase() === ref.sheet?.toLowerCase(),
									)
								: si
							if (refSi !== sheetIndex) continue
							if (
								row >= ref.start.row &&
								row <= ref.end.row &&
								col >= ref.start.col &&
								col <= ref.end.col
							) {
								feedsInto.push(`${s.name}!${colLabel(c)}${r + 1}`)
							}
						}
					}
				}
			}
		}

		return {
			ref: cellRef,
			formula: cell?.formula ?? null,
			dependsOn,
			feedsInto,
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
		return [...this.wb.definedNames.keys()]
	}
}

// --- Helpers ---

import { indexToColumn } from '@ascend/core'
import type { FormulaNode } from '@ascend/formulas'

function colLabel(col: number): string {
	return indexToColumn(col)
}

function parseFullRef(
	cellRef: string,
	workbook: Workbook,
): { sheetName: string; row: number; col: number } {
	const bang = cellRef.indexOf('!')
	if (bang !== -1) {
		const sheetName = cellRef.substring(0, bang).replace(/^'|'$/g, '')
		const parsed = parseRange(cellRef.substring(bang + 1))
		return { sheetName, row: parsed.start.row, col: parsed.start.col }
	}
	const firstSheet = workbook.sheets[0]
	const sheetName = firstSheet ? firstSheet.name : 'Sheet1'
	const parsed = parseRange(cellRef)
	return { sheetName, row: parsed.start.row, col: parsed.start.col }
}

function findVolatileCalls(node: FormulaNode, volatileFns: Set<string>): string[] {
	const found: string[] = []
	walkAst(node, (n) => {
		if (n.type === 'function' && volatileFns.has(n.name.toUpperCase())) {
			found.push(n.name.toUpperCase())
		}
	})
	return found
}

function walkAst(node: FormulaNode, visitor: (n: FormulaNode) => void): void {
	visitor(node)
	switch (node.type) {
		case 'binary':
			walkAst(node.left, visitor)
			walkAst(node.right, visitor)
			break
		case 'unary':
			walkAst(node.operand, visitor)
			break
		case 'function':
			for (const arg of node.args) walkAst(arg, visitor)
			break
		case 'array':
			for (const row of node.rows) {
				for (const cell of row) walkAst(cell, visitor)
			}
			break
	}
}
