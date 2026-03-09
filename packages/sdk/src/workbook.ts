import { writeFile } from 'node:fs/promises'
import { createWorkbook, parseRange, type RangeRef, type Workbook } from '@ascend/core'
import {
	applyOperations,
	type CalcContext,
	defaultCalcContext,
	diffWorkbooks,
	recalculate,
} from '@ascend/engine'
import { readCsv, writeCsv } from '@ascend/io-csv'
import { type PreservationCapsule, writeXlsx } from '@ascend/io-xlsx'
import {
	type CompatibilityReport,
	type CsvDialect,
	emptyReport,
	type Operation,
} from '@ascend/schema'
import { check as verifyCheck, lint as verifyLint } from '@ascend/verify'
import { buildWorkbookLoadInfo, openWorkbookSource } from './load.ts'
import { WorkbookReadView } from './read-view.ts'
import type {
	ApplyResult,
	CheckIssue,
	CheckResult,
	LintResult,
	LintWarning,
	RecalcResult,
} from './types.ts'

function cloneWorkbook(source: Workbook): Workbook {
	return source.clone()
}

export class AscendWorkbook extends WorkbookReadView {
	private readonly caps: PreservationCapsule[]
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
		super(workbook, report, loadInfo)
		this.caps = capsules
		this.originalBytes = originalBytes
		this.dirty = false
	}

	static async open(
		pathOrBytes: string | Uint8Array,
		options?: { mode?: 'full' | 'metadata-only' | 'values'; sheets?: readonly string[] },
	): Promise<AscendWorkbook> {
		const loaded = await openWorkbookSource(pathOrBytes, options)
		return new AscendWorkbook(
			loaded.workbook,
			[...loaded.capsules],
			loaded.report,
			loaded.loadInfo,
			loaded.originalBytes,
		)
	}

	static create(): AscendWorkbook {
		const wb = createWorkbook()
		wb.addSheet('Sheet1')
		return new AscendWorkbook(
			wb,
			[],
			emptyReport('ascend'),
			buildWorkbookLoadInfo({
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
			buildWorkbookLoadInfo({
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
