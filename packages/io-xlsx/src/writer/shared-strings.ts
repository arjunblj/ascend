import type { Workbook } from '@ascend/core'
import type { CellValue, RichTextRun } from '@ascend/schema'
import { escapeXml } from '../xml.ts'
import { ChunkedStringBuilder } from './chunked-string-builder.ts'
import type { DynamicArrayMetadataEntry } from './metadata.ts'

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

export interface SharedStringTable {
	getIndex(value: CellValue): number | undefined
	toXml(): string
	readonly count: number
	readonly facts: WorkbookWriteFacts
}

export interface WorkbookWriteFacts {
	readonly hasStringCells: boolean
	readonly dynamicArrayMetadataEntries: readonly DynamicArrayMetadataEntry[]
}

export class IncrementalSharedStringTable {
	private readonly entries: CellValue[] = []
	private readonly lookup = new Map<string, number>()
	private _count = 0
	readonly facts: WorkbookWriteFacts

	constructor(existingEntries: readonly CellValue[], facts: WorkbookWriteFacts) {
		this.facts = facts
		for (let i = 0; i < existingEntries.length; i++) {
			const entry = existingEntries[i]
			if (!entry) continue
			const key = makeKey(entry)
			if (key !== undefined && !this.lookup.has(key)) {
				this.lookup.set(key, this.entries.length)
				this.entries.push(entry)
			}
		}
	}

	getIndex(value: CellValue): number | undefined {
		const key = makeKey(value)
		if (key === undefined) return undefined
		let idx = this.lookup.get(key)
		if (idx === undefined) {
			idx = this.entries.length
			this.lookup.set(key, idx)
			this.entries.push(value)
		}
		this._count++
		return idx
	}

	toXml(): string {
		const builder = new ChunkedStringBuilder()
		builder.push(XML_HEADER)
		builder.push(`<sst xmlns="${NS}" count="${this._count}" uniqueCount="${this.entries.length}">`)
		for (const entry of this.entries) {
			builder.push(entryXml(entry))
		}
		builder.push('</sst>')
		return builder.toString()
	}

	get count(): number {
		return this._count
	}

	get uniqueCount(): number {
		return this.entries.length
	}
}

export function scanWorkbookWriteFactsFast(workbook: Workbook): WorkbookWriteFacts {
	let hasStringCells = false
	let dynamicArrayMetadataEntries: readonly DynamicArrayMetadataEntry[] = []

	for (const sheet of workbook.sheets) {
		if (hasStringCells && dynamicArrayMetadataEntries.length > 0) break
		for (const [, , cell] of sheet.cells.iterate()) {
			if (!hasStringCells) {
				const key = makeKey(cell.value)
				if (key !== undefined) hasStringCells = true
			}
			if (dynamicArrayMetadataEntries.length === 0) {
				const binding = cell.formulaInfo
				if (binding) {
					if (binding.kind === 'dynamicArray') {
						dynamicArrayMetadataEntries = [
							{ metadataIndex: 1, collapsed: binding.collapsed ?? false },
						]
					} else if (binding.kind === 'spill' && binding.isAnchor) {
						dynamicArrayMetadataEntries = [{ metadataIndex: 1, collapsed: false }]
					}
				}
			}
			if (hasStringCells && dynamicArrayMetadataEntries.length > 0) break
		}
	}
	return { hasStringCells, dynamicArrayMetadataEntries }
}

function makeKey(value: CellValue): string | undefined {
	if (value.kind === 'string') return `s:${value.value}`
	if (value.kind === 'richText') return richTextKey(value.runs)
	return undefined
}

function richTextKey(runs: readonly import('@ascend/schema').RichTextRun[]): string {
	let key = 'r:'
	for (const run of runs) {
		key += `${run.text}\x01${run.bold ? 1 : 0}\x01${run.italic ? 1 : 0}\x01${run.underline ? 1 : 0}\x01${run.strikethrough ? 1 : 0}\x01${run.fontSize ?? ''}\x01${run.color ?? ''}\x01${run.fontName ?? ''}\x02`
	}
	return key
}

function entryXml(value: CellValue): string {
	if (value.kind === 'string') {
		return `<si><t>${escapeXml(value.value)}</t></si>`
	}
	if (value.kind === 'richText') {
		const runs = value.runs.map(runXml).join('')
		return `<si>${runs}</si>`
	}
	return '<si><t/></si>'
}

function runXml(run: RichTextRun): string {
	const rPr = runPropsXml(run)
	const rPrEl = rPr ? `<rPr>${rPr}</rPr>` : ''
	return `<r>${rPrEl}<t>${escapeXml(run.text)}</t></r>`
}

function runPropsXml(run: RichTextRun): string {
	const parts: string[] = []
	if (run.bold) parts.push('<b/>')
	if (run.italic) parts.push('<i/>')
	if (run.underline) parts.push('<u/>')
	if (run.strikethrough) parts.push('<strike/>')
	if (run.fontSize !== undefined) parts.push(`<sz val="${run.fontSize}"/>`)
	if (run.color) parts.push(`<color rgb="${escapeXml(run.color)}"/>`)
	if (run.fontName) parts.push(`<rFont val="${escapeXml(run.fontName)}"/>`)
	return parts.join('')
}
