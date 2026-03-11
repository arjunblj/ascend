import type { Workbook } from '@ascend/core'
import type { CellValue, RichTextRun } from '@ascend/schema'
import { escapeXml } from '../xml.ts'
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

export function buildSharedStrings(
	workbook: Workbook,
	existingEntries: readonly CellValue[] = [],
): SharedStringTable {
	const entries: CellValue[] = [...existingEntries]
	const lookup = new Map<string, number>()
	for (let index = 0; index < entries.length; index++) {
		const entry = entries[index]
		if (!entry) continue
		const key = makeKey(entry)
		if (key !== undefined && !lookup.has(key)) lookup.set(key, index)
	}
	let count = 0
	let hasStringCells = false
	let dynamicArrayMetadataEntries: readonly DynamicArrayMetadataEntry[] = []

	for (const sheet of workbook.sheets) {
		for (const [, , cell] of sheet.cells.iterate()) {
			const key = makeKey(cell.value)
			if (key !== undefined) {
				hasStringCells = true
				count += 1
				if (!lookup.has(key)) {
					lookup.set(key, entries.length)
					entries.push(cell.value)
				}
			}
			const binding = cell.formulaInfo
			if (dynamicArrayMetadataEntries.length === 0 && binding) {
				if (binding.kind === 'dynamicArray') {
					dynamicArrayMetadataEntries = [
						{ metadataIndex: 1, collapsed: binding.collapsed ?? false },
					]
				} else if (binding.kind === 'spill' && binding.isAnchor) {
					dynamicArrayMetadataEntries = [{ metadataIndex: 1, collapsed: false }]
				}
			}
		}
	}
	const facts = { hasStringCells, dynamicArrayMetadataEntries }

	return {
		getIndex(value: CellValue): number | undefined {
			const key = makeKey(value)
			return key !== undefined ? lookup.get(key) : undefined
		},
		toXml(): string {
			const parts: string[] = [
				XML_HEADER,
				`<sst xmlns="${NS}" count="${count}" uniqueCount="${entries.length}">`,
			]
			for (const entry of entries) {
				parts.push(entryXml(entry))
			}
			parts.push('</sst>')
			return parts.join('')
		},
		count,
		facts,
	}
}

export function scanWorkbookWriteFacts(workbook: Workbook): WorkbookWriteFacts {
	let hasStringCells = false
	let dynamicArrayMetadataEntries: readonly DynamicArrayMetadataEntry[] = []
	for (const sheet of workbook.sheets) {
		for (const [, , cell] of sheet.cells.iterate()) {
			if (!hasStringCells && (cell.value.kind === 'string' || cell.value.kind === 'richText')) {
				hasStringCells = true
			}
			const binding = cell.formulaInfo
			if (dynamicArrayMetadataEntries.length === 0 && binding) {
				if (binding.kind === 'dynamicArray') {
					dynamicArrayMetadataEntries = [
						{ metadataIndex: 1, collapsed: binding.collapsed ?? false },
					]
				} else if (binding.kind === 'spill' && binding.isAnchor) {
					dynamicArrayMetadataEntries = [{ metadataIndex: 1, collapsed: false }]
				}
			}
			if (hasStringCells && dynamicArrayMetadataEntries.length > 0) {
				return { hasStringCells, dynamicArrayMetadataEntries }
			}
		}
	}
	return { hasStringCells, dynamicArrayMetadataEntries }
}

function makeKey(value: CellValue): string | undefined {
	if (value.kind === 'string') return `s:${value.value}`
	if (value.kind === 'richText') return `r:${JSON.stringify(value.runs)}`
	return undefined
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
