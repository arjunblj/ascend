import type { CellValue, RichTextRun } from '@ascend/schema'
import { asArray, attr, numAttr, parseXml, type XmlNode } from '../xml.ts'

export interface SharedStringResolver {
	readonly count: number
	get(index: number): CellValue | undefined
}

export function parseSharedStrings(
	xml: string,
	options: {
		readonly normalize?: (value: CellValue) => CellValue
	} = {},
): SharedStringResolver {
	return createEagerSharedStrings(xml, options.normalize)
}

export function emptySharedStrings(): SharedStringResolver {
	return {
		count: 0,
		get(): CellValue | undefined {
			return undefined
		},
	}
}

function createEagerSharedStrings(
	xml: string,
	normalize?: (value: CellValue) => CellValue,
): SharedStringResolver {
	const doc = parseXml(xml)
	const sst = doc.sst as XmlNode | undefined
	if (!sst) return emptySharedStrings()

	const entries: CellValue[] = []

	for (const si of asArray<XmlNode>(sst.si as XmlNode | XmlNode[])) {
		entries.push(normalize ? normalize(parseSharedStringNode(si)) : parseSharedStringNode(si))
	}

	return {
		count: entries.length,
		get(index: number): CellValue | undefined {
			return entries[index]
		},
	}
}

function parseSharedStringNode(si: XmlNode): CellValue {
	if (si.t !== undefined) {
		return { kind: 'string', value: String(si.t) }
	}
	if (si.r !== undefined) {
		return parseRichText(si)
	}
	return { kind: 'string', value: '' }
}

function parseRichText(si: XmlNode): CellValue {
	const runs: RichTextRun[] = []

	for (const r of asArray<XmlNode>(si.r as XmlNode | XmlNode[])) {
		const text = r.t !== undefined ? String(r.t) : ''
		const rPr = r.rPr as XmlNode | undefined

		if (rPr && typeof rPr === 'object') {
			const run: RichTextRun = {
				text,
				...(rPr.b !== undefined ? { bold: true } : {}),
				...(rPr.i !== undefined ? { italic: true } : {}),
				...(rPr.u !== undefined ? { underline: true } : {}),
				...(rPr.strike !== undefined ? { strikethrough: true } : {}),
				...parseFontProps(rPr),
			}
			runs.push(run)
		} else {
			runs.push({ text })
		}
	}

	const first = runs[0]
	if (
		runs.length === 1 &&
		first &&
		!first.bold &&
		!first.italic &&
		!first.underline &&
		!first.strikethrough &&
		!first.fontName &&
		!first.fontSize &&
		!first.color
	) {
		return { kind: 'string', value: first.text }
	}

	return { kind: 'richText', runs }
}

function parseFontProps(rPr: XmlNode): Pick<RichTextRun, 'fontName' | 'fontSize' | 'color'> {
	const result: Pick<RichTextRun, 'fontName' | 'fontSize' | 'color'> = {}

	const rFont = rPr.rFont
	if (typeof rFont === 'object' && rFont !== null) {
		const name = attr(rFont as XmlNode, 'val')
		if (name) (result as Record<string, unknown>).fontName = name
	}

	const sz = rPr.sz
	if (typeof sz === 'object' && sz !== null) {
		const size = numAttr(sz as XmlNode, 'val')
		if (size !== undefined) (result as Record<string, unknown>).fontSize = size
	}

	const color = rPr.color
	if (typeof color === 'object' && color !== null) {
		const rgb = attr(color as XmlNode, 'rgb')
		if (rgb) (result as Record<string, unknown>).color = rgb
	}

	return result
}
