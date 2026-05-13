import type { Operation } from '@ascend/schema'
import {
	type ApplyAndRecalcResult,
	AscendWorkbook,
	auditLossPolicy,
	type CellSelector,
	type CompactCellInfo,
	type CompactRangeInfo,
	type FormulaInfo,
	type PreviewResult,
	type SheetInspectInfo,
	WorkbookDocument,
	type WorkbookInfo,
	type WorkbookLoadOptions,
} from '@ascend/sdk'
import type { CommandIntent, CommandPreview, OpenWorkbook } from '../runtime/types.ts'

export interface SaveTarget {
	readonly path?: string
	readonly saveCopy?: boolean
}

export interface SaveResult {
	readonly ok: boolean
	readonly path: string
	readonly message: string
	readonly dirty: boolean
	readonly current: boolean
}

export class WorkbookSessionController {
	private workbook: AscendWorkbook | null = null
	private document: WorkbookDocument | null = null
	private path: string | null = null
	private dirty = false
	private dirtyVersion = 0
	private workbookVersion = 0

	async open(path: string, options: WorkbookLoadOptions = {}): Promise<OpenWorkbook> {
		if (isReadOnlyDocumentLoad(options)) {
			this.document = await WorkbookDocument.open(path, options)
			this.workbook = null
		} else {
			this.workbook = await AscendWorkbook.open(path, options)
			this.document = null
		}
		this.workbookVersion += 1
		this.path = path
		this.dirty = false
		this.dirtyVersion = 0
		const info = this.inspect()
		if (!info) throw new Error('Workbook opened without inspect metadata.')
		const readOnly = info.load.isPartial
		return {
			id: path,
			path,
			name: path.split(/[\\/]/).pop() ?? path,
			info,
			readOnly,
			protectedReview: hasProtectedReviewSignals(info),
			dirty: false,
		}
	}

	createEmpty(): OpenWorkbook {
		this.workbook = AscendWorkbook.create()
		this.document = null
		this.workbookVersion += 1
		this.path = null
		this.dirty = true
		this.dirtyVersion += 1
		return {
			id: 'untitled',
			path: null,
			name: 'Book1',
			info: this.workbook.inspect(),
			readOnly: false,
			protectedReview: false,
			dirty: true,
		}
	}

	requireWorkbook(): AscendWorkbook {
		if (!this.workbook) throw new Error('No workbook is open.')
		return this.workbook
	}

	inspect(): WorkbookInfo | null {
		return this.workbook?.inspect() ?? this.document?.inspect() ?? null
	}

	inspectSheet(name: string): SheetInspectInfo | undefined {
		return this.workbook?.inspectSheet(name) ?? this.document?.inspectSheet(name)
	}

	readRangeCompact(
		sheetName: string,
		range: string,
		options?: {
			readonly includeRefs?: boolean
			readonly omitEmpty?: boolean
			readonly flatValues?: boolean
		},
	): CompactRangeInfo | undefined {
		return (
			this.workbook?.readRangeCompact(sheetName, range, options) ??
			this.document?.readRangeCompact(sheetName, range, options)
		)
	}

	formula(cellRef: CellSelector): FormulaInfo | undefined {
		return this.workbook?.formula(cellRef) ?? this.document?.formula(cellRef)
	}

	cellCompact(sheetName: string, ref: CellSelector): CompactCellInfo | undefined {
		return (
			this.workbook?.sheet(sheetName)?.cellCompact(ref) ??
			this.document?.sheet(sheetName)?.cellCompact(ref)
		)
	}

	previewOperations(intent: CommandIntent, operations: readonly Operation[]): CommandPreview {
		return { intent, operations, warnings: [] }
	}

	preview(operations: readonly Operation[]): PreviewResult {
		return this.requireWorkbook().preview(operations)
	}

	applyAndRecalc(operations: readonly Operation[]): ApplyAndRecalcResult | null {
		if (operations.length === 0) return null
		const result = this.requireWorkbook().applyAndRecalc(operations)
		if (result.apply.errors.length === 0) {
			this.dirty = true
			this.dirtyVersion += 1
		}
		if (result.apply.errors.length > 0 || (result.recalc?.errors.length ?? 0) > 0) return result
		return result
	}

	async save(target?: SaveTarget): Promise<SaveResult> {
		const targetPath = target?.path ?? this.path
		if (!targetPath) {
			return {
				ok: false,
				path: '',
				message: 'Use Save As for an unnamed workbook.',
				dirty: this.dirty,
				current: true,
			}
		}
		const workbook = this.requireWorkbook()
		const saveDirtyVersion = this.dirtyVersion
		const saveWorkbookVersion = this.workbookVersion
		await workbook.save(targetPath)
		const current = this.workbook === workbook && this.workbookVersion === saveWorkbookVersion
		if (current && !target?.saveCopy) {
			this.path = targetPath
			if (this.dirtyVersion === saveDirtyVersion) this.dirty = false
		}
		return {
			ok: true,
			path: targetPath,
			message: `Saved ${targetPath}`,
			dirty: current ? this.dirty : true,
			current,
		}
	}

	isDirty(): boolean {
		return this.dirty
	}

	currentPath(): string | null {
		return this.path
	}
}

function isReadOnlyDocumentLoad(options: WorkbookLoadOptions): boolean {
	return options.mode === 'values' && options.maxRows !== undefined
}

function hasProtectedReviewSignals(info: WorkbookInfo): boolean {
	return (
		!auditLossPolicy(info.compatibility.features).ok ||
		info.externalReferenceCount > 0 ||
		info.pivotTableCount > 0 ||
		info.chartCount > 0
	)
}
