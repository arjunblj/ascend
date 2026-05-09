import type { Operation } from '@ascend/schema'
import {
	type ApplyAndRecalcResult,
	AscendWorkbook,
	auditLossPolicy,
	type PreviewResult,
	type WorkbookInfo,
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
	private path: string | null = null
	private dirty = false
	private dirtyVersion = 0
	private workbookVersion = 0

	async open(path: string): Promise<OpenWorkbook> {
		this.workbook = await AscendWorkbook.open(path)
		this.workbookVersion += 1
		this.path = path
		this.dirty = false
		this.dirtyVersion = 0
		const info = this.workbook.inspect()
		return {
			id: path,
			path,
			name: path.split(/[\\/]/).pop() ?? path,
			info,
			readOnly: false,
			protectedReview: hasProtectedReviewSignals(info),
			dirty: false,
		}
	}

	createEmpty(): OpenWorkbook {
		this.workbook = AscendWorkbook.create()
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
		return this.workbook?.inspect() ?? null
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

function hasProtectedReviewSignals(info: WorkbookInfo): boolean {
	return (
		!auditLossPolicy(info.compatibility.features).ok ||
		info.externalReferenceCount > 0 ||
		info.pivotTableCount > 0 ||
		info.chartCount > 0
	)
}
