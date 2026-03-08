import type { CellValue, RichTextRun } from '@ascend/schema'
import { asArray, attr, numAttr, parseXml, type XmlNode } from '../xml.ts'

export function parseSharedStrings(xml: string): CellValue[] {
	const doc = parseXml(xml)
	const sst = doc.sst as XmlNode | undefined
	if (!sst) return []

	const entries: CellValue[] = []

	for (const si of asArray<XmlNode>(sst.si as XmlNode | XmlNode[])) {
		if (si.t !== undefined) {
			entries.push({ kind: 'string', value: String(si.t) })
		} else if (si.r !== undefined) {
			entries.push(parseRichText(si))
		} else {
			entries.push({ kind: 'string', value: '' })
		}
	}

	return entries
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
